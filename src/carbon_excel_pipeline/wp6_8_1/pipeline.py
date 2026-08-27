"""Capability-driven single-Run end-to-end orchestrator for WP6-8.1."""

from __future__ import annotations

import csv
import json
import shutil
from decimal import Decimal
from pathlib import Path
from typing import Any, Mapping

from carbon_excel_pipeline.activity.day5_pipeline import run_day5_upstream_rebuild
from carbon_excel_pipeline.capability.pipeline import run_wp6_2_capability_detection
from carbon_excel_pipeline.calculation.day7_pipeline import run_day7_calculation_and_lineage
from carbon_excel_pipeline.cleaning.scope_filter import run_day3_scope_and_cleaning
from carbon_excel_pipeline.errors import PipelineUserError
from carbon_excel_pipeline.factors.day6_pipeline import run_day6_factors_and_matching
from carbon_excel_pipeline.io.excel_importer import inspect_excel_to_run, sha256_file
from carbon_excel_pipeline.standardization.standardizer import run_day4_standardization
from carbon_excel_pipeline.wp6_3.pipeline import run_wp6_3_historical_reproduction
from carbon_excel_pipeline.wp6_8.pipeline import run_live_delivery
from carbon_excel_pipeline.wp6_8_1.independent_live import validate_live_canonical
from carbon_excel_pipeline.wp6_8_1.pcs_adapter import adapt_pcs_current_run
from carbon_excel_pipeline.wp6_8_1.router import (
    ACTIVITY_DIRECT,
    ACTIVITY_PCS,
    FACTOR_SIMULATION,
    decide_processing_routes,
)
from carbon_excel_pipeline.wp6_8_4.attribute_enrichment import (
    enrich_canonical_with_attributes,
    load_attribute_rows,
)
from carbon_excel_pipeline.wp6_8_4.business_units import (
    additional_direct_mass_rows,
    additional_pcs_sheet_rows,
    detect_business_units,
    unit_label_from_sheet,
)
from carbon_excel_pipeline.wp6_8_4.file_roles import classify_workbook_role, reconcile_roles
from carbon_excel_pipeline.wp6_8_4.input_set import (
    ROLE_ATTRIBUTE,
    ROLE_LEDGER,
    ROLE_UNKNOWN,
    input_set_sha256,
)
from carbon_excel_pipeline.wp6_8_5.cell_scope import apply_cell_scope_to_run
from carbon_excel_pipeline.wp6_8_5.current_run import persist_current_run
from carbon_excel_pipeline.wp6_8_5.ledger_reference import extract_cell_ledger_evidence
from carbon_excel_pipeline.wp6_8_5.public_factor import (
    apply_public_cell_factor,
    reconcile_with_ledger,
)
from carbon_excel_pipeline.wp6_8_6.record_ids import (
    RECORD_ID_SCHEMA_VERSION,
    RecordIDSchemaError,
    assign_record_ids,
    load_record_id_config,
    propagate_record_ids_in_run,
    remap_record_ids_in_rows,
)
from carbon_excel_pipeline.io.header_detector import load_alias_config


PROJECT_ROOT = Path(__file__).resolve().parents[3]
FORBIDDEN_LIVE_FRAGMENTS = (
    "wp6-3\\runs",
    "wp6-3/runs",
    "wp6-4\\runs",
    "wp6-4/runs",
    "wp6-5\\runs",
    "wp6-5/runs",
    "wp6-6\\runs",
    "wp6-6/runs",
    "wp6-7\\runs",
    "wp6-7/runs",
    "wp6-8\\runs",
    "wp6-8/runs",
)
STAGE_ORDER = (
    "UPLOADED",
    "RECOGNIZED",
    "CAPABILITY_DETECTED",
    "ROUTED",
    "PROCESSING",
    "CALCULATED",
    "VALIDATED",
    "ANALYZED",
    "PACKAGED",
    "COMPLETED",
)


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row:
            if key not in seen:
                seen.add(key)
                fields.append(key)
    if not fields:
        fields = ["Record_ID"]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _assert_not_formal_wp6_run(path: Path, *, current_run: Path | None = None) -> None:
    resolved = path.expanduser().resolve()
    if current_run is not None:
        try:
            resolved.relative_to(current_run.resolve())
            return
        except ValueError:
            pass
    text = str(resolved).lower()
    for fragment in FORBIDDEN_LIVE_FRAGMENTS:
        if fragment.lower() in text:
            raise PipelineUserError(
                stage="E2E_ORCHESTRATOR",
                error_code="FORMAL_WP6_RUN_USED_AS_LIVE_INPUT",
                message_cn="Live Run 不得读取 WP6-3～WP6-8 已有正式计算结果作为输入。",
                source_location=str(resolved),
                original_value=fragment,
                rule="旧正式 Run 只允许作为 Expected Regression Comparison。",
                impact="阻断本次一键核算，避免伪完成。",
                fix_suggestion="只提供原始 Excel 和当前 Run 内产物。",
            )


def _receipt_sha(run_dir: Path) -> str:
    receipt = _load_json(run_dir / "01_import" / "file_receipt_report.json")
    return str(receipt.get("source_sha256", "")).upper()


def _can_reuse_run(run_dir: Path | None, source_sha: str) -> bool:
    if run_dir is None or not run_dir.is_dir():
        return False
    receipt = run_dir / "01_import" / "file_receipt_report.json"
    if not receipt.is_file():
        return False
    if _receipt_sha(run_dir) != source_sha:
        return False
    if (run_dir / "e2e_run_summary.json").is_file():
        return False
    calculation = run_dir / "04_calculation"
    if calculation.is_dir() and any(calculation.iterdir()):
        return False
    return (run_dir / "01_import" / "recognized_records.json").is_file()


def _stage_status(run_dir: Path, **updates: Any) -> dict[str, Any]:
    path = run_dir / "pipeline_stage_status.json"
    payload = _load_json(path) if path.is_file() else {
        "schema_version": "WP6_8_1_PIPELINE_STAGE_STATUS_V1",
        "Run_ID": run_dir.name,
        "stages": {name: "NOT_RUN" for name in STAGE_ORDER},
    }
    payload.update(updates)
    _write_json(path, payload)
    return payload


def _set_stage(run_dir: Path, name: str, status: str) -> None:
    payload = _stage_status(run_dir)
    payload["stages"][name] = status
    payload["current_stage"] = name
    _write_json(run_dir / "pipeline_stage_status.json", payload)


def _stamp_additional_canonical(rows: list[dict[str, Any]], *, run_id: str) -> list[dict[str, Any]]:
    """Keep detected-unit Factor/Boundary fields. Do not stamp 二部 SHA-gated policies."""

    enriched: list[dict[str, Any]] = []
    for row in rows:
        blocking = str(row.get("Blocking_Codes") or "")
        emission = str(row.get("Emission_kgCO2e") or "").strip()
        enriched.append(
            {
                **row,
                "Run_ID": run_id,
                "Activity_Unit": row.get("Activity_Unit") or "kg/year",
                "Activity_Ready": "TRUE" if row.get("Activity_Data_kg") else "FALSE",
                "Factor_Ready": "TRUE" if row.get("EF_Value") else "FALSE",
                "Emission_Ready": "TRUE" if emission else "FALSE",
                "Boundary_Ready": "FALSE" if "BOUNDARY_POLICY_NOT_AVAILABLE" in blocking or blocking else row.get("Boundary_Ready") or "TRUE",
                "Simulation_Flag": "TRUE",
                "Production_Eligible": "FALSE",
            }
        )
    return enriched


def _fill_missing_business_units(rows: list[dict[str, Any]]) -> None:
    for row in rows:
        if str(row.get("Business_Unit") or "").strip():
            continue
        sheet = str(row.get("Source_Sheet") or "").strip()
        if sheet:
            row["Business_Unit"] = unit_label_from_sheet(sheet)


def _stamp_attribute_governance(rows: list[dict[str, Any]]) -> None:
    fields = {
        "Chemistry": "CHEMISTRY_MISSING",
        "Supplier": "SUPPLIER_MISSING",
        "Project": "PROJECT_MISSING",
        "Model": "MODEL_MISSING",
    }
    for row in rows:
        codes = [
            code
            for field, code in fields.items()
            if not str(row.get(field) or row.get({"Supplier": "Supplier_Name", "Project": "Project_Code", "Model": "Cell_Model"}.get(field, "")) or "").strip()
        ]
        row["Governance_Issue_Codes"] = "|".join(codes)
        if codes:
            row["Governance_QC"] = "WARNING"


def _enrich_direct_canonical(
    rows: list[dict[str, Any]],
    *,
    run_id: str,
    route: Mapping[str, Any],
) -> list[dict[str, Any]]:
    enriched: list[dict[str, Any]] = []
    for row in rows:
        warning = str(row.get("Warning_Codes") or "")
        emission = str(row.get("Emission_kgCO2e") or "").strip()
        blocking = str(row.get("Blocking_Codes") or "").strip()
        calculation_qc = "BLOCKED" if blocking or not emission else "PASS"
        governance_qc = "WARNING" if "MISSING" in warning.upper() else (
            "PASS" if str(row.get("QC_Status")) == "PASS" else str(row.get("QC_Status") or "WARNING")
        )
        enriched.append(
            {
                **row,
                "Run_ID": run_id,
                "Activity_Unit": row.get("Activity_Unit") or "kg/year",
                "Activity_Ready": "TRUE" if row.get("Activity_Data_kg") else "FALSE",
                "Factor_Ready": "TRUE" if row.get("EF_Value") else "FALSE",
                "Emission_Ready": "TRUE" if emission else "FALSE",
                "Boundary_Ready": "FALSE" if blocking else "TRUE",
                "Calculation_QC": calculation_qc,
                "Governance_QC": governance_qc,
                "EF_Usage": route.get("Factor_Usage") or row.get("EF_Usage") or "SOURCE_EMBEDDED_FACTOR",
                "Factor_Route": route.get("Factor_Route"),
                "Factor_Policy_ID": route.get("Factor_Policy_ID"),
                "Boundary_Policy": route.get("Boundary_Policy"),
                "Simulation_Flag": "TRUE",
                "Production_Eligible": "FALSE",
            }
        )
    return enriched


def _run_direct_mass(
    run_dir: Path,
    *,
    route: Mapping[str, Any],
    project_root: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    output_dir = run_dir / "04_calculation"
    policy_path = project_root / str(route["Boundary_Config_Path"])
    _assert_not_formal_wp6_run(policy_path, current_run=run_dir)
    result = run_wp6_3_historical_reproduction(
        run_dir,
        output_root=run_dir / "_unused_wp6_3_root",
        policy_config_path=policy_path,
        run_id=run_dir.name,
        output_dir=output_dir,
    )
    rows = _enrich_direct_canonical(
        _read_csv(output_dir / "2024_canonical_results.csv"),
        run_id=run_dir.name,
        route=route,
    )
    historical = _read_csv(output_dir / "2024_historical_validation.csv")
    return rows, historical, result


def _run_pcs_current_run(
    run_dir: Path,
    *,
    route: Mapping[str, Any],
    project_root: Path,
    private_paths: Mapping[str, Path],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    if route.get("Factor_Route") != FACTOR_SIMULATION:
        raise PipelineUserError(
            stage="E2E_ORCHESTRATOR",
            error_code="PCS_FACTOR_POLICY_NOT_CONTROLLED",
            message_cn="PCS×单重路径没有匹配到受控历史模拟因子 Policy，不能自动套用 1.250000。",
            source_location="Factor_Route",
            original_value=route.get("Factor_Route"),
            rule="HISTORICAL_SIMULATION_FACTOR 必须由受控 Policy 选择。",
            impact="活动数据可以保留，排放结果不能生成。",
            fix_suggestion="为该文件建立受控因子 Policy，或提供嵌入 EF。",
        )
    for key, path in private_paths.items():
        _assert_not_formal_wp6_run(path, current_run=run_dir)
    scope_path = project_root / str(route["Boundary_Config_Path"])
    factor_path = project_root / str(route["Factor_Config_Path"])
    day3 = run_day3_scope_and_cleaning(
        run_dir,
        scope_config_path=scope_path,
        unit_config_path=project_root / "config" / "cleaning" / "unit_mappings.json",
    )
    day4 = run_day4_standardization(
        run_dir,
        profile_config_path=project_root / "config" / "profiles" / "public_synthetic_profile.json",
        contract_path=project_root / "config" / "standardization" / "standard_31_contract.json",
        mapping_config_path=project_root / "config" / "mapping" / "public_synthetic_mapping_v1.json",
        private_id_baseline_path=private_paths["private_id_baseline"],
        private_mapping_workbook_path=private_paths["private_mapping_workbook"],
    )
    day5 = run_day5_upstream_rebuild(
        run_dir,
        profile_config_path=project_root / "config" / "profiles" / "public_synthetic_profile.json",
        standard_contract_path=project_root / "config" / "standardization" / "standard_31_contract.json",
        activity_contract_path=project_root / "config" / "activity" / "activity_36_contract.json",
        third_party_contract_path=project_root / "config" / "activity" / "third_party_20_contract.json",
        quality_config_path=project_root / "config" / "qc" / "day5_quality_rules.json",
        interface_open_items_path=project_root / "config" / "open_items" / "wp3_interface_open_items.json",
        standard_baseline_path=private_paths["standard_baseline"],
        activity_baseline_path=private_paths["activity_baseline"],
        third_party_baseline_path=private_paths["third_party_baseline"],
    )
    day6 = run_day6_factors_and_matching(
        run_dir,
        profile_config_path=project_root / "config" / "profiles" / "public_synthetic_profile.json",
        activity_contract_path=project_root / "config" / "activity" / "activity_36_contract.json",
        external_contract_path=project_root / "config" / "factors" / "external_8_contract.json",
        d1_contract_path=project_root / "config" / "factors" / "wp5_d1_45_contract.json",
        factor_config_path=factor_path,
        d2_contract_path=project_root / "config" / "matching" / "wp5_d2_57_contract.json",
        route_contract_path=project_root / "config" / "matching" / "wp5_d3_route_36_contract.json",
        historical_simulation=True,
    )
    day7 = run_day7_calculation_and_lineage(
        run_dir,
        profile_config_path=project_root / "config" / "profiles" / "public_synthetic_profile.json",
        calculation_rules_path=project_root / "config" / "calculation" / "day7_calculation_rules.json",
        activity_contract_path=project_root / "config" / "activity" / "activity_36_contract.json",
        d1_contract_path=project_root / "config" / "factors" / "wp5_d1_45_contract.json",
        d2_contract_path=project_root / "config" / "matching" / "wp5_d2_57_contract.json",
        d3_contract_path=project_root / "config" / "matching" / "wp5_d3_route_36_contract.json",
        d4_contract_path=project_root / "config" / "calculation" / "wp5_d4_48_contract.json",
        d5_contract_path=project_root / "config" / "calculation" / "wp5_d5_56_contract.json",
        frozen_lineage_contract_path=project_root / "config" / "lineage" / "wp5_frozen_32_contract.json",
        extended_lineage_contract_path=project_root / "config" / "lineage" / "demo_extended_lineage_contract.json",
    )
    rows = adapt_pcs_current_run(run_dir, route_decision=dict(route))
    output_dir = run_dir / "04_calculation"
    output_dir.mkdir(parents=True, exist_ok=True)
    _write_csv(output_dir / "canonical_results.csv", rows)
    _write_json(
        output_dir / "pcs_processing_summary.json",
        {
            "run_id": run_dir.name,
            "record_count": len(rows),
            "day3": day3.get("status"),
            "day4": day4.get("status"),
            "day5": day5.get("status"),
            "day6": day6.get("status"),
            "day7": day7.get("status"),
            "second_run_created": False,
        },
    )
    return rows, [], {
        "status": day7.get("status"),
        "record_count": len(rows),
        "second_run_created": False,
    }


def _totals(rows: list[dict[str, Any]]) -> dict[str, str]:
    activity = sum((Decimal(str(row["Activity_Data_kg"])) for row in rows if row.get("Activity_Data_kg")), Decimal("0"))
    emission_kg = sum((Decimal(str(row["Emission_kgCO2e"])) for row in rows if row.get("Emission_kgCO2e")), Decimal("0"))
    emission_t = sum((Decimal(str(row["Emission_tCO2e"])) for row in rows if row.get("Emission_tCO2e")), Decimal("0"))
    return {
        "activity_kg": format(activity, "f"),
        "emission_kgco2e": format(emission_kg, "f"),
        "emission_tco2e": format(emission_t, "f"),
    }


def _source_sheet_orders(run_dir: Path) -> dict[str, int]:
    path = run_dir / "01_import" / "sheet_inventory.json"
    if not path.is_file():
        return {}
    payload = _load_json(path)
    return {
        str(item.get("sheet_name") or ""): int(item.get("sheet_index") or 0)
        for item in payload
        if item.get("sheet_name")
    }


def run_end_to_end_pipeline(
    source_path: Path,
    *,
    run_root: Path,
    project_root: Path = PROJECT_ROOT,
    existing_run_dir: Path | None = None,
    private_paths: Mapping[str, Path] | None = None,
    requested_activity_route: str | None = None,
    extra_source_paths: list[Path] | None = None,
    file_roles: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    """Upload one primary Excel (and optional attribute files) into one Current Run."""

    source = source_path.expanduser().resolve()
    _assert_not_formal_wp6_run(source)
    source_sha = sha256_file(source)
    alias = project_root / "config" / "import" / "field_aliases.json"
    reused = _can_reuse_run(existing_run_dir, source_sha)
    if reused:
        run_dir = existing_run_dir.expanduser().resolve()
        inspect_result = {
            "status": "REUSED",
            "run_directory": str(run_dir),
            "run_id": run_dir.name,
        }
    else:
        inspect_result = inspect_excel_to_run(
            source, run_root=run_root, alias_config_path=alias
        )
        run_dir = Path(str(inspect_result["run_directory"])).resolve()
    cell_scope = apply_cell_scope_to_run(run_dir)
    extra_paths = [Path(path).expanduser().resolve() for path in (extra_source_paths or []) if path]
    for extra in extra_paths:
        _assert_not_formal_wp6_run(extra)
        destination = run_dir / "00_input_copy" / extra.name
        destination.parent.mkdir(parents=True, exist_ok=True)
        if extra.resolve() != destination.resolve():
            shutil.copy2(extra, destination)
    alias_config = load_alias_config(alias)
    input_files = [
        {"name": source.name, "sha256": source_sha, "role": "主核算数据", "path": str(source)}
    ]
    for extra in extra_paths:
        classified = classify_workbook_role(extra, alias_config)
        classified["sha256"] = sha256_file(extra)
        input_files.append(classified)
    if file_roles:
        for item in input_files:
            override = file_roles.get(item.get("name") or "") or file_roles.get(item.get("sha256") or "")
            if override:
                item["role"] = override
                item["suggested"] = False
    input_files = reconcile_roles(input_files)
    set_sha = input_set_sha256(input_files)
    _set_stage(run_dir, "UPLOADED", "PASS")
    _set_stage(run_dir, "RECOGNIZED", inspect_result.get("status", "PASS"))

    capability = run_wp6_2_capability_detection(run_dir)
    capability_summary = _load_json(run_dir / "02_capability" / "capability_summary.json")
    recognition_summary = _load_json(run_dir / "01_import" / "recognition_summary.json")
    _set_stage(run_dir, "CAPABILITY_DETECTED", capability_summary.get("status", "PASS"))

    route = decide_processing_routes(
        capability=capability_summary,
        input_sha256=source_sha,
        requested_activity_route=requested_activity_route,
    )
    route_dir = run_dir / "03_route"
    _write_json(route_dir / "route_decision.json", route)
    _write_json(run_dir / "route_decision.json", route)
    _set_stage(run_dir, "ROUTED", route["status"])

    frozen_references = []
    if private_paths:
        frozen_references = [str(path) for path in private_paths.values()]
    provenance = {
        "schema_version": "WP6_8_1_PROVENANCE_V1",
        "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
        "Run_ID": run_dir.name,
        "Input_File": source.name,
        "Input_SHA256": source_sha,
        "Input_Set_SHA256": set_sha,
        "Input_Files": input_files,
        "live_input_paths": [str(source), *[str(path) for path in extra_paths], str(run_dir)],
        "frozen_reference_paths": frozen_references,
        "historical_formal_runs_used_as_live_input": [],
        "second_business_run_created": False,
        "year_used_as_router": False,
        "cross_year_factor_analysis": "NOT_AVAILABLE_FOR_SINGLE_INPUT_RUN",
    }
    _write_json(run_dir / "provenance.json", provenance)

    ledger_evidence: dict[str, Any] | None = None
    ledger_warnings: list[str] = [
        f"{item.get('name')}: 未识别辅助文件用途，已保留但不影响主核算。"
        for item in input_files
        if item.get("role") == ROLE_UNKNOWN
    ]
    for item in input_files:
        if item.get("role") != ROLE_LEDGER or not item.get("path"):
            continue
        evidence = extract_cell_ledger_evidence(Path(str(item["path"])))
        if evidence.get("status") == "PASS" and ledger_evidence is None:
            ledger_evidence = evidence
        else:
            ledger_warnings.append(
                f"{item.get('name')}: 未提取到完整电芯类别因子证据，作为辅助文件保留。"
            )
    if ledger_evidence or ledger_warnings:
        _write_json(
            run_dir / "03_route" / "historical_ledger_evidence.json",
            ledger_evidence or {"status": "WARNING", "warnings": ledger_warnings},
        )

    canonical_rows: list[dict[str, Any]] = []
    historical_rows: list[dict[str, Any]] = []
    processing: dict[str, Any] = {"status": route["status"]}
    overall = route["status"]
    if route["status"] == "ROUTED" and route.get("Processor") == "DIRECT_MASS_WP6_3":
        _set_stage(run_dir, "PROCESSING", "RUNNING")
        canonical_rows, historical_rows, processing = _run_direct_mass(
            run_dir, route=route, project_root=project_root
        )
        _set_stage(run_dir, "PROCESSING", "PASS")
        _set_stage(run_dir, "CALCULATED", processing.get("status", "PASS"))
    elif route["status"] == "ROUTED" and route.get("Processor") == "PCS_DAY3_TO_DAY7":
        if not private_paths:
            raise PipelineUserError(
                stage="E2E_ORCHESTRATOR",
                error_code="PCS_PRIVATE_PATHS_MISSING",
                message_cn="PCS 路径需要 WP2/WP3 冻结基线才能在当前 Run 内复用 Day3–Day7。",
                source_location="private_paths",
                original_value="",
                rule="不得为了省事改调 run_day8_full_pipeline 另建 Run。",
                impact="无法在同一 Current Run 完成 2025 核算。",
                fix_suggestion="传入与 Streamlit 相同的只读 WP2/WP3 路径。",
            )
        _set_stage(run_dir, "PROCESSING", "RUNNING")
        canonical_rows, historical_rows, processing = _run_pcs_current_run(
            run_dir, route=route, project_root=project_root, private_paths=private_paths
        )
        _set_stage(run_dir, "PROCESSING", "PASS")
        _set_stage(run_dir, "CALCULATED", processing.get("status", "PASS"))
    else:
        _set_stage(run_dir, "PROCESSING", route["status"])
        _set_stage(run_dir, "CALCULATED", "NOT_RUN")

    frozen_provisional_ids = {
        str(row.get("Record_ID") or "") for row in canonical_rows if row.get("Record_ID")
    }
    detected_units: list[str] = []
    attribute_note = ""
    if canonical_rows and route.get("Processor") == "DIRECT_MASS_WP6_3":
        policy = _load_json(project_root / str(route["Boundary_Config_Path"]))
        extra_rows, extra_hist, detected_units = additional_direct_mass_rows(
            run_dir,
            existing_canonical=canonical_rows,
            policy=policy,
            run_id=run_dir.name,
        )
        canonical_rows.extend(_stamp_additional_canonical(extra_rows, run_id=run_dir.name))
        historical_rows.extend(extra_hist)
    elif canonical_rows and route.get("Processor") == "PCS_DAY3_TO_DAY7":
        extra_rows, detected_units = additional_pcs_sheet_rows(
            run_dir,
            existing_canonical=canonical_rows,
            run_id=run_dir.name,
            input_sha256=source_sha,
            source_name=source.name,
        )
        canonical_rows.extend(_stamp_additional_canonical(extra_rows, run_id=run_dir.name))
    _fill_missing_business_units(canonical_rows)
    detected_units = list(
        dict.fromkeys([*detected_units, *detect_business_units(canonical_rows)])
    )
    attribute_sources = [
        Path(item["path"])
        for item in input_files
        if item.get("role") == ROLE_ATTRIBUTE and item.get("path")
    ]
    if attribute_sources and canonical_rows:
        attribute_rows: list[dict[str, Any]] = []
        for path in attribute_sources:
            attribute_rows.extend(load_attribute_rows(path, alias_config))
        canonical_rows, enrich_summary = enrich_canonical_with_attributes(
            canonical_rows, attribute_rows
        )
        attribute_note = enrich_summary.get("message") or ""
        _write_json(run_dir / "03_route" / "attribute_enrichment.json", enrich_summary)
    _stamp_attribute_governance(canonical_rows)

    record_id_summary: dict[str, Any] = {
        "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
        "status": "NOT_RUN",
    }
    record_id_propagation: dict[str, Any] = {
        "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
        "status": "NOT_RUN",
    }
    frozen_record_ids: set[str] = set()
    if canonical_rows:
        try:
            record_id_config = load_record_id_config(
                project_root / "config" / "wp6" / "record_id_schema_v2.json"
            )
            canonical_rows, record_id_mapping, record_id_summary = assign_record_ids(
                canonical_rows,
                config=record_id_config,
                source_sheet_orders=_source_sheet_orders(run_dir),
            )
        except RecordIDSchemaError as error:
            _set_stage(run_dir, "RECORD_ID_ASSIGNED", "BLOCKED")
            raise PipelineUserError(
                stage="RECORD_ID_ASSIGNED",
                error_code="RID_V2_GENERATION_BLOCKED",
                message_cn=str(error),
                source_location="Canonical Record_ID",
                original_value="",
                rule="RID_V2 必须非空、唯一、格式合法并按稳定物理源顺序编号。",
                impact="当前 Run 被阻断，禁止以随机字符或旧编号继续。",
                fix_suggestion="补齐受控事业部、供应商或物料代码映射后重新运行。",
            ) from error
        historical_rows = remap_record_ids_in_rows(historical_rows, record_id_mapping)
        frozen_record_ids = {
            record_id_mapping[old_id]
            for old_id in frozen_provisional_ids
            if old_id in record_id_mapping
        }
        record_id_propagation = propagate_record_ids_in_run(run_dir, record_id_mapping)
        _write_json(
            run_dir / "03_route" / "record_id_schema.json",
            {**record_id_summary, "propagation": record_id_propagation},
        )
        _set_stage(run_dir, "RECORD_ID_ASSIGNED", "PASS")

    canonical_rows, public_factor_summary = apply_public_cell_factor(
        canonical_rows, ledger_evidence
    )
    _write_json(run_dir / "03_route" / "public_cell_factor_application.json", public_factor_summary)
    ledger_reconciliation = reconcile_with_ledger(canonical_rows, ledger_evidence)
    if ledger_evidence:
        _write_json(run_dir / "05_validation" / "historical_ledger_reconciliation.json", ledger_reconciliation)

    effective_route = dict(route)
    if public_factor_summary.get("status") == "PASS":
        effective_route.update(
            {
                "Factor_Route": "HISTORICAL_SIMULATION_FACTOR",
                "Factor_Policy_ID": public_factor_summary.get("policy_id"),
                "Factor_Value": public_factor_summary.get("factor_value"),
                "Factor_Unit": public_factor_summary.get("factor_unit"),
                "Factor_Source": ledger_evidence.get("Source_File") if ledger_evidence else None,
                "Factor_Usage": "HISTORICAL_SIMULATION",
                "Factor_Route_Reason": "清册电芯类别证据、Purchase_Category=电芯及已确认事业部范围共同满足。",
                "Factor_Ready": True,
                "Boundary_Policy": public_factor_summary.get("boundary_policy_id"),
                "Boundary_Ready": True,
                "Boundary_Policy_Reason": "当前记录属于已识别的合成一部或二部电芯业务范围。",
                "Emission_Ready": True,
                "Simulation_Flag": True,
                "Production_Eligible": False,
            }
        )
        _write_json(run_dir / "03_route" / "effective_route_decision.json", effective_route)

    validation_rows: list[dict[str, Any]] = []
    independent_status = "NOT_RUN"
    frozen_rows = [
        row for row in canonical_rows if str(row.get("Record_ID") or "") in frozen_record_ids
    ] or canonical_rows
    if canonical_rows:
        calculable = [row for row in canonical_rows if str(row.get("Emission_kgCO2e") or "").strip()]
        validation_rows = validate_live_canonical(calculable) if calculable else []
        by_id = {row.get("Record_ID"): row for row in validation_rows}
        for row in canonical_rows:
            matched = by_id.get(row.get("Record_ID"), {})
            row["Overall_Validation_Status"] = matched.get(
                "Overall_Validation_Status", row.get("Overall_Status") or "NOT_RUN"
            )
        frozen_validation = [
            row for row in validation_rows if row.get("Record_ID") in {item.get("Record_ID") for item in frozen_rows}
        ]
        statuses = {row.get("Overall_Validation_Status") for row in frozen_validation}
        if statuses == {"INDEPENDENT_CALCULATION_PASS"}:
            independent_status = "INDEPENDENT_CALCULATION_PASS"
        elif not statuses:
            independent_status = "NOT_RUN"
        elif "INDEPENDENT_VALIDATION_FAIL" in statuses:
            independent_status = "INDEPENDENT_VALIDATION_FAIL"
        else:
            independent_status = next(iter(statuses)) if len(statuses) == 1 else "PARTIAL_RESULT"
        validation_dir = run_dir / "05_validation"
        _write_csv(validation_dir / "independent_validation.csv", validation_rows)
        _write_json(
            validation_dir / "independent_validation_summary.json",
            {
                "status": independent_status,
                "record_count": len(validation_rows),
                "source": "wp6_5.independent live arithmetic",
                "formal_wp6_5_run_read": False,
            },
        )
        _set_stage(run_dir, "VALIDATED", independent_status)
    else:
        _set_stage(run_dir, "VALIDATED", "NOT_RUN")

    delivery: dict[str, Any] = {}
    if canonical_rows:
        delivery = run_live_delivery(
            run_dir=run_dir,
            canonical_rows=canonical_rows,
            route_decision=effective_route,
            validation_rows=validation_rows,
            historical_rows=historical_rows,
            input_file=source.name,
            input_sha256=source_sha,
            recognition_summary=recognition_summary,
            capability_summary=capability_summary,
            independent_status=independent_status,
        )
        _set_stage(run_dir, "ANALYZED", delivery.get("status", "PASS"))
        _set_stage(run_dir, "PACKAGED", delivery.get("status", "PASS"))
        overall = delivery.get("status", overall)
        _write_csv(run_dir / "08_download" / "canonical_results.csv", canonical_rows)
        _write_csv(run_dir / "04_calculation" / "canonical_results.csv", canonical_rows)
    else:
        _set_stage(run_dir, "ANALYZED", "NOT_RUN")
        _set_stage(run_dir, "PACKAGED", "NOT_RUN")
        overall = route["status"]

    totals = _totals(canonical_rows)
    frozen_totals = _totals(frozen_rows)
    summary = {
        "schema_version": "WP6_8_1_E2E_RUN_SUMMARY_V1",
        "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
        "Run_ID": run_dir.name,
        "Status": overall,
        "Input_File": source.name,
        "Input_SHA256": source_sha,
        "Input_Set_SHA256": set_sha,
        "Input_Files": input_files,
        "Recognition_Status": recognition_summary.get("recognition_status"),
        "Selected_Sheet": recognition_summary.get("best_candidate_sheet"),
        "Header_Row": recognition_summary.get("best_candidate_header_row"),
        "Activity_Route": route.get("Activity_Route"),
        "Factor_Route": effective_route.get("Factor_Route"),
        "Factor_Policy_ID": effective_route.get("Factor_Policy_ID"),
        "Factor_Value": effective_route.get("Factor_Value"),
        "Factor_Source": effective_route.get("Factor_Source"),
        "Factor_Usage": effective_route.get("Factor_Usage"),
        "Simulation_Flag": True,
        "Production_Eligible": False,
        "Boundary_Policy": effective_route.get("Boundary_Policy"),
        "Boundary_Ready": effective_route.get("Boundary_Ready"),
        "Record_Count": len(canonical_rows),
        "Totals": totals,
        "Detected_Business_Units": detected_units,
        "Frozen_Baseline": {
            "record_count": len(frozen_rows),
            "totals": frozen_totals,
        },
        "Attribute_Match_Note": attribute_note,
        "Record_ID_Validation": record_id_summary,
        "Record_ID_Propagation": record_id_propagation,
        "Cell_Scope": cell_scope,
        "Public_Cell_Factor": public_factor_summary,
        "Historical_Ledger_Evidence": ledger_evidence or {"status": "NOT_PROVIDED"},
        "Historical_Ledger_Reconciliation": ledger_reconciliation,
        "Auxiliary_File_Warnings": ledger_warnings,
        "Independent_Validation_Status": independent_status,
        "Cross_Year_Factor_Analysis": "NOT_AVAILABLE_FOR_SINGLE_INPUT_RUN",
        "Second_Business_Run_Created": False,
        "Historical_Formal_Runs_Used_As_Live_Input": [],
        "Year_Used_As_Router": False,
        "Inspect_Reused": reused,
        "Delivery_Directory": delivery.get("output_directory"),
        "Provenance_File": str(run_dir / "provenance.json"),
    }
    _write_json(run_dir / "e2e_run_summary.json", summary)
    _set_stage(run_dir, "COMPLETED", overall)
    persist_current_run(run_dir)
    return {
        "status": overall,
        "run_id": run_dir.name,
        "run_directory": str(run_dir),
        "input_sha256": source_sha,
        "input_set_sha256": set_sha,
        "route_decision": effective_route,
        "record_count": len(canonical_rows),
        "record_id_schema_version": RECORD_ID_SCHEMA_VERSION,
        "totals": totals,
        "frozen_baseline": summary["Frozen_Baseline"],
        "detected_business_units": detected_units,
        "independent_validation_status": independent_status,
        "cross_year_factor_analysis": "NOT_AVAILABLE_FOR_SINGLE_INPUT_RUN",
        "second_business_run_created": False,
        "e2e_run_summary": summary,
        "delivery": delivery,
        "capability": capability,
        "inspect": inspect_result,
    }
