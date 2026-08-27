"""WP6-4 2025 strict regression and different-scope reconciliation."""

from __future__ import annotations

import csv
import hashlib
import json
from collections import Counter
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.errors import PipelineUserError

from .enterprise import extract_enterprise_cell_scope
from .quality import (
    assess_quality_layers,
    ef_audit,
    forward_fill_audit,
    preferred_precision_status,
)


CURRENT_FILES = {
    "recognized": "01_import/recognized_records.json",
    "capability": "02_capability/capability_summary.json",
    "scope": "02_scope_filter/day3_scope_summary.json",
    "standard": "03_standardized/day4_standard_31_fields.csv",
    "activity": "05_activity/day5_activity_36_fields.csv",
    "result": "10_output/day7_d5_end_to_end_56_fields.csv",
    "lineage": "10_output/day7_wp5_frozen_lineage_32_fields.csv",
    "lineage_qa": "10_output/day7_lineage_qa.json",
    "summary": "10_output/day7_summary.json",
}


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    columns = fields or (list(rows[0]) if rows else [])
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _boolean(value: bool) -> str:
    return "TRUE" if value else "FALSE"


def _decimal_equal(left: Any, right: Any) -> bool:
    try:
        return Decimal(str(left)) == Decimal(str(right))
    except Exception:
        return False


def _error(code: str, message: str, value: Any, suggestion: str) -> PipelineUserError:
    return PipelineUserError(
        stage="WP6-4",
        error_code=code,
        message_cn=message,
        source_location="WP6-4 input",
        original_value=value,
        rule="Current Run、Frozen Evidence 与历史清册必须保持不同验证角色。",
        impact="WP6-4 正式严格回归停止。",
        fix_suggestion=suggestion,
    )


def _require_current_run(run: Path) -> dict[str, Path]:
    paths = {name: run / relative for name, relative in CURRENT_FILES.items()}
    missing = [str(path) for path in paths.values() if not path.is_file()]
    if missing:
        raise _error("CURRENT_RUN_INCOMPLETE", "2025 Current Run 缺少正式阶段产物。", missing, "先使用现有共享流水线完成 Day3—Day7。")
    capability = _load_json(paths["capability"])
    if int(capability.get("pcs_weight_derived_count", 0)) <= 0:
        raise _error("PCS_WEIGHT_CAPABILITY_MISSING", "上游能力产物不支持 PCS_WEIGHT_DERIVED。", capability, "检查 WP6-2 正式能力产物。")
    return paths


def _build_boundary_audit(recognized: dict[str, Any], current_activity: list[dict[str, str]]) -> list[dict[str, Any]]:
    records = recognized["records"]
    cell_rows = []
    for record in records:
        category = str(record.get("values", {}).get("Purchase_Category") or "")
        if category.split(".", 1)[0].strip() == "电芯":
            cell_rows.append(record)
    return [
        {
            "Stage": "BUSINESS_UNIT_SHEET_CONTEXT",
            "Input_Count": len(records),
            "Selected_Count": len(records),
            "Excluded_Count": 0,
            "Filter_Rule": f"Recognized source sheet = {recognized['sheet_name']}",
        },
        {
            "Stage": "PURCHASE_CATEGORY",
            "Input_Count": len(records),
            "Selected_Count": len(cell_rows),
            "Excluded_Count": len(records) - len(cell_rows),
            "Filter_Rule": "Purchase_Category root exact = 电芯",
        },
        {
            "Stage": "SYNTHETIC_CELL_BOUNDARY",
            "Input_Count": len(cell_rows),
            "Selected_Count": len(current_activity),
            "Excluded_Count": len(cell_rows) - len(current_activity),
            "Filter_Rule": "Existing PUBLIC_SYNTHETIC_DAY3_SCOPE_V1 business rule",
        },
    ]


def _build_report(summary: dict[str, Any]) -> str:
    strict = summary["strict_regression"]
    scope = summary["enterprise_scope_reconciliation"]
    qc = summary["quality"]
    ef = summary["ef_audit"]
    return "\n".join(
        [
            "# WP6-4 2025严格回归与范围勾稽报告",
            "",
            f"> 正式状态：{summary['status']}",
            f"> Run：{summary['run_id']}",
            "",
            "## 2025 二部 SYNA 严格试点回归",
            "",
            f"- 记录：{strict['record_count']} 条",
            f"- 状态：`{strict['status']}`",
            f"- Activity：{strict['activity_total_kg']} kg/year",
            f"- Emission：{strict['unrounded_total_kgco2e']} kgCO2e/year",
            f"- 汇总后六位：{strict['rounded_aggregate_total']} kgCO2e/year",
            f"- 逐行六位合计：{strict['sum_of_row_rounded']} kgCO2e/year",
            f"- 舍入勾稽：{strict['rounding_reconciliation']} kgCO2e/year",
            "",
            "## 质量与边界",
            "",
            f"- Calculation QC：{qc['Calculation_QC_Counts']}",
            f"- Governance QC：{qc['Governance_QC_Counts']}",
            f"- Overall：{qc['Overall_Status_Counts']}",
            f"- Boundary Ready：{qc['Boundary_Ready_Counts']}",
            f"- Emission Ready：{qc['Emission_Ready_Counts']}",
            "",
            "## EF 审计",
            "",
            f"- EF 种类数：{ef['EF_Unique_Count']}",
            f"- 最小/最大：{ef['EF_Min']} / {ef['EF_Max']}",
            f"- Activity 加权 EF：{ef['Activity_Weighted_EF']}",
            "",
            "## 企业历史范围勾稽",
            "",
            "企业历史为全电芯类别；Demo 为合成二部 SYNA 合成示例。两者核算范围不同，不要求总量一致。",
            "",
            f"- Activity Coverage：{scope['activity_coverage_percent']}%",
            f"- Emission Coverage：{scope['emission_coverage_percent']}%",
            f"- Scope Comparison：`{scope['comparison_type']}`",
            "",
            "## 阶段边界",
            "",
            "本阶段未执行 WP6-5 独立第二套复算、因子影响分析、管理 Dashboard 或建议清册输出。",
            "",
        ]
    )


def run_wp6_4_validation(
    current_run_dir: Path,
    *,
    frozen_result_path: Path,
    frozen_lineage_path: Path,
    frozen_activity_path: Path,
    enterprise_workbook_path: Path,
    wp6_3_run_dir: Path,
    raw_input_path: Path,
    output_root: Path,
    strict_contract_path: Path,
    run_id: str | None = None,
) -> dict[str, Any]:
    current = current_run_dir.expanduser().resolve()
    paths = _require_current_run(current)
    sources = {
        "raw_input": raw_input_path.expanduser().resolve(),
        "frozen_result": frozen_result_path.expanduser().resolve(),
        "frozen_lineage": frozen_lineage_path.expanduser().resolve(),
        "frozen_activity": frozen_activity_path.expanduser().resolve(),
        "enterprise_workbook": enterprise_workbook_path.expanduser().resolve(),
        "wp6_3_canonical": wp6_3_run_dir.expanduser().resolve() / "2024_canonical_results.csv",
        "wp6_3_validation": wp6_3_run_dir.expanduser().resolve() / "2024_historical_validation.csv",
    }
    missing = [str(path) for path in sources.values() if not path.is_file()]
    if missing:
        raise _error("FORMAL_SOURCE_MISSING", "WP6-4 正式输入不完整。", missing, "恢复正式只读证据路径后重试。")
    before_hashes = {name: _sha256(path) for name, path in sources.items()}
    contract = _load_json(strict_contract_path.expanduser().resolve())
    recognized = _load_json(paths["recognized"])
    current_summary = _load_json(paths["summary"])
    current_standard = _read_csv(paths["standard"])
    current_activity = _read_csv(paths["activity"])
    current_result = _read_csv(paths["result"])
    current_lineage = _read_csv(paths["lineage"])
    frozen_result = _read_csv(sources["frozen_result"])
    frozen_lineage = _read_csv(sources["frozen_lineage"])
    frozen_activity = _read_csv(sources["frozen_activity"])

    def unique(rows: list[dict[str, str]], label: str) -> dict[str, dict[str, str]]:
        output = {row["Record_ID"]: row for row in rows}
        if len(output) != len(rows):
            raise _error("DUPLICATE_RECORD_ID", f"{label} 存在重复 Record_ID。", len(rows) - len(output), "修复证据唯一性。")
        return output

    current_by_id = unique(current_result, "Current Result")
    frozen_by_id = unique(frozen_result, "Frozen Result")
    activity_by_id = unique(current_activity, "Current Activity")
    frozen_activity_by_id = unique(frozen_activity, "Frozen Activity")
    standard_by_id = unique(current_standard, "Current Standard")
    current_lineage_by_id = unique(current_lineage, "Current Lineage")
    frozen_lineage_by_id = unique(frozen_lineage, "Frozen Lineage")
    current_ids = [row["Record_ID"] for row in current_result]
    frozen_ids = [row["Record_ID"] for row in frozen_result]

    activity_rows: list[dict[str, Any]] = []
    ef_rows: list[dict[str, Any]] = []
    emission_rows: list[dict[str, Any]] = []
    qc_rows: list[dict[str, Any]] = []
    lineage_rows: list[dict[str, Any]] = []
    strict_rows: list[dict[str, Any]] = []
    canonical_rows: list[dict[str, Any]] = []
    for index, record_id in enumerate(current_ids, start=1):
        current_row = current_by_id[record_id]
        frozen_row = frozen_by_id.get(record_id, {})
        activity = activity_by_id[record_id]
        frozen_activity_row = frozen_activity_by_id.get(record_id, {})
        standard = standard_by_id[record_id]
        quality = assess_quality_layers(activity=activity, result=current_row, standard=standard)
        activity_match = _decimal_equal(current_row.get("Activity_Data_Normalized_kg"), frozen_row.get("Activity_Data_Normalized_kg"))
        ef_match = _decimal_equal(current_row.get("EF_Value_Normalized_kgCO2e_per_kg"), frozen_row.get("EF_Value_Normalized_kgCO2e_per_kg"))
        raw_emission_match = _decimal_equal(current_row.get("Raw_Emission_kgCO2e"), frozen_row.get("Raw_Emission_kgCO2e"))
        display_emission_match = _decimal_equal(current_row.get("Emission_kgCO2e"), frozen_row.get("Emission_kgCO2e"))
        governance_match = activity.get("QC_Status") == frozen_activity_row.get("QC_Status")
        lineage_match = all(
            activity.get(field) == frozen_activity_row.get(field)
            for field in ("Source_File", "Source_Sheet", "Source_Row")
        ) and record_id in current_lineage_by_id and record_id in frozen_lineage_by_id
        sequence_match = index <= len(frozen_ids) and frozen_ids[index - 1] == record_id
        record_pass = all((sequence_match, activity_match, ef_match, raw_emission_match, display_emission_match, governance_match, lineage_match))
        activity_rows.append({
            "Record_ID": record_id,
            "Current_Activity_kg": current_row.get("Activity_Data_Normalized_kg", ""),
            "Frozen_Activity_kg": frozen_row.get("Activity_Data_Normalized_kg", ""),
            "Difference_kg": format(Decimal(current_row["Activity_Data_Normalized_kg"]) - Decimal(frozen_row["Activity_Data_Normalized_kg"]), "f") if frozen_row else "",
            "Validation_Status": "PASS" if activity_match else "FAIL",
        })
        ef_rows.append({
            "Record_ID": record_id,
            "Current_EF": current_row.get("EF_Value_Normalized_kgCO2e_per_kg", ""),
            "Frozen_EF": frozen_row.get("EF_Value_Normalized_kgCO2e_per_kg", ""),
            "Difference": format(Decimal(current_row["EF_Value_Normalized_kgCO2e_per_kg"]) - Decimal(frozen_row["EF_Value_Normalized_kgCO2e_per_kg"]), "f") if frozen_row else "",
            "Validation_Status": "PASS" if ef_match else "FAIL",
        })
        emission_rows.append({
            "Record_ID": record_id,
            "Current_Emission_kgCO2e": current_row.get("Raw_Emission_kgCO2e", ""),
            "Frozen_Emission_kgCO2e": frozen_row.get("Raw_Emission_kgCO2e", ""),
            "Difference_kgCO2e": format(Decimal(current_row["Raw_Emission_kgCO2e"]) - Decimal(frozen_row["Raw_Emission_kgCO2e"]), "f") if frozen_row else "",
            "Validation_Status": "PASS" if raw_emission_match else "FAIL",
        })
        qc_rows.append({"Record_ID": record_id, "Frozen_Governance_QC": frozen_activity_row.get("QC_Status", ""), **quality, "Governance_QC_Match": _boolean(governance_match)})
        lineage_rows.append({
            "Record_ID": record_id,
            "Current_Source_File": activity.get("Source_File", ""),
            "Frozen_Source_File": frozen_activity_row.get("Source_File", ""),
            "Current_Source_Sheet": activity.get("Source_Sheet", ""),
            "Frozen_Source_Sheet": frozen_activity_row.get("Source_Sheet", ""),
            "Current_Source_Row": activity.get("Source_Row", ""),
            "Frozen_Source_Row": frozen_activity_row.get("Source_Row", ""),
            "Validation_Status": "PASS" if lineage_match else "FAIL",
        })
        strict_rows.append({
            "Record_ID": record_id,
            "Sequence_Match": _boolean(sequence_match),
            "Activity_Match": _boolean(activity_match),
            "EF_Match": _boolean(ef_match),
            "Raw_Emission_Match": _boolean(raw_emission_match),
            "Display_Emission_Match": _boolean(display_emission_match),
            "Governance_QC_Match": _boolean(governance_match),
            "Lineage_Match": _boolean(lineage_match),
            "Validation_Status": "PASS" if record_pass else "FAIL",
        })
        canonical_rows.append({
            "Record_ID": record_id,
            "Year": standard.get("Year", ""),
            "Source_File": activity.get("Source_File", ""),
            "Source_Sheet": activity.get("Source_Sheet", ""),
            "Source_Row": activity.get("Source_Row", ""),
            "Business_Unit": standard.get("Business_Unit", ""),
            "Purchase_Category": standard.get("Activity_Category", ""),
            "Product_Description": standard.get("Product_Description", ""),
            "Activity_Data_kg": current_row.get("Activity_Data_Normalized_kg", ""),
            "Activity_Unit": current_row.get("Activity_Unit_Normalized", ""),
            "Activity_Method": "PCS_WEIGHT_DERIVED",
            "EF_Value": current_row.get("EF_Value_Normalized_kgCO2e_per_kg", ""),
            "EF_Unit": current_row.get("EF_Unit_Normalized", ""),
            "Emission_kgCO2e": current_row.get("Raw_Emission_kgCO2e", ""),
            **quality,
            "Simulation_Flag": current_row.get("Simulation_Flag", ""),
            "Production_Eligible": current_row.get("Production_Eligible", ""),
        })

    boundary_audit = _build_boundary_audit(recognized, current_activity)
    ef_summary = ef_audit(current_result)
    quality_summary = {
        "Calculation_QC_Counts": dict(Counter(row["Calculation_QC"] for row in qc_rows)),
        "Governance_QC_Counts": dict(Counter(row["Governance_QC"] for row in qc_rows)),
        "Overall_Status_Counts": dict(Counter(row["Overall_Status"] for row in qc_rows)),
        "Boundary_Ready_Counts": dict(Counter(_boolean(bool(row["Boundary_Ready"])) for row in qc_rows)),
        "Emission_Ready_Counts": dict(Counter(_boolean(bool(row["Emission_Ready"])) for row in qc_rows)),
    }
    expected_rows = int(contract["upstream_baseline"]["rows"])
    calculation_expected = contract["calculation_baseline"]
    strict_checks = {
        "Record_Count_Match": len(current_ids) == len(frozen_ids) == expected_rows,
        "Record_ID_Match": current_ids == frozen_ids,
        "Activity_Match": all(row["Validation_Status"] == "PASS" for row in activity_rows),
        "EF_Match": all(row["Validation_Status"] == "PASS" for row in ef_rows),
        "Emission_Match": all(row["Validation_Status"] == "PASS" for row in emission_rows),
        "QC_Match": all(row["Governance_QC_Match"] == "TRUE" for row in qc_rows),
        "Lineage_Match": all(row["Validation_Status"] == "PASS" for row in lineage_rows),
        "Aggregate_Match": current_summary["raw_total_emission_kgco2e"] == calculation_expected["unrounded_total_kgco2e_per_year"],
        "Rounding_Match": (
            current_summary["official_six_decimal_total"] == calculation_expected["official_six_decimal_total"]
            and current_summary["row_six_decimal_sum"] == calculation_expected["row_six_decimal_sum"]
            and current_summary["rounding_reconciliation_difference"] == calculation_expected["rounding_reconciliation_difference"]
        ),
        "Calculation_QC_Pass": quality_summary["Calculation_QC_Counts"] == {"PASS": expected_rows},
        "Boundary_Ready": quality_summary["Boundary_Ready_Counts"] == {"TRUE": expected_rows},
    }
    strict_status = "STRICT_REGRESSION_PASS" if all(strict_checks.values()) else "STRICT_REGRESSION_FAIL"

    enterprise = extract_enterprise_cell_scope(sources["enterprise_workbook"])
    demo_activity = Decimal(current_summary["activity_total_kg"])
    demo_emission_kg = Decimal(current_summary["raw_total_emission_kgco2e"])
    enterprise_activity = Decimal(enterprise["activity_kg"])
    enterprise_emission_t = Decimal(enterprise["historical_emission_tco2e"])
    activity_coverage = demo_activity / enterprise_activity
    emission_coverage = demo_emission_kg / (enterprise_emission_t * Decimal("1000"))
    scope_reconciliation = {
        "comparison_type": "DIFFERENT_SCOPE_RECONCILIATION_NOT_STRICT_REGRESSION",
        "enterprise_scope": enterprise,
        "demo_scope": {
            "scope": "SYNTHETIC_DIVISION_2_SYNA_PILOT",
            "records": len(current_ids),
            "activity_kg": format(demo_activity, "f"),
            "ef_kgco2e_per_kg": ef_summary["Activity_Weighted_EF"],
            "emission_kgco2e": format(demo_emission_kg, "f"),
            "emission_tco2e": format(demo_emission_kg / Decimal("1000"), "f"),
        },
        "activity_coverage": format(activity_coverage, "f"),
        "activity_coverage_percent": format(activity_coverage * Decimal("100"), ".6f"),
        "emission_coverage": format(emission_coverage, "f"),
        "emission_coverage_percent": format(emission_coverage * Decimal("100"), ".6f"),
        "same_scope": False,
        "total_equality_required": False,
        "different_scope_is_strict_regression_failure": False,
    }
    scope_rows = [
        {"Dimension": "Business Unit", "Enterprise_Scope": "ALL / NOT DISAGGREGATED IN EXTRACTED TOTAL", "Demo_Scope": "合成二部", "Difference_Status": "DIFFERENT"},
        {"Dimension": "Material", "Enterprise_Scope": "全电芯类别", "Demo_Scope": "SYNA 电芯严格试点", "Difference_Status": "DIFFERENT"},
        {"Dimension": "Supplier / Product", "Enterprise_Scope": "NOT AVAILABLE AT EXTRACTED TOTAL", "Demo_Scope": "SYNA / synthetic records", "Difference_Status": "PARTIALLY_AVAILABLE"},
        {"Dimension": "Record Scope", "Enterprise_Scope": "AGGREGATED CATEGORY TOTAL", "Demo_Scope": f"{len(current_ids)} RECORDS", "Difference_Status": "DIFFERENT"},
        {"Dimension": "Activity", "Enterprise_Scope": enterprise["activity_kg"], "Demo_Scope": format(demo_activity, "f"), "Difference_Status": "COVERAGE_ONLY"},
        {"Dimension": "Emission_tCO2e", "Enterprise_Scope": enterprise["historical_emission_tco2e"], "Demo_Scope": format(demo_emission_kg / Decimal("1000"), "f"), "Difference_Status": "COVERAGE_ONLY"},
    ]

    wp63_rows = _read_csv(sources["wp6_3_canonical"])
    wp63_validation = _read_csv(sources["wp6_3_validation"])
    preferred_counts = Counter(
        preferred_precision_status(row.get("Validation_Status", ""))
        for row in wp63_validation
    )
    wp63_forward = {
        "source_run": str(wp6_3_run_dir.expanduser().resolve()),
        "business_result_recalculated": False,
        "formal_run_modified": False,
        "controlled_forward_fill": forward_fill_audit(wp63_rows),
        "preferred_precision_terminology": "FORMULA_CACHE_PRECISION_DIFFERENCE",
        "preferred_validation_status_counts": dict(preferred_counts),
    }

    current_run_id = run_id or (
        "WP6-4-" + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ") + "-" + recognized["input_fingerprint"][:8]
    )
    output_dir = output_root.expanduser().resolve() / current_run_id
    if output_dir.exists():
        raise _error("OUTPUT_RUN_ALREADY_EXISTS", "WP6-4 Run 目录已存在。", str(output_dir), "使用新的 Run_ID，禁止覆盖正式证据。")
    output_dir.mkdir(parents=True)

    _write_csv(output_dir / "2025_boundary_filter.csv", boundary_audit)
    _write_csv(output_dir / "2025_strict_regression.csv", strict_rows)
    _write_csv(output_dir / "2025_activity_regression.csv", activity_rows)
    _write_csv(output_dir / "2025_ef_regression.csv", ef_rows)
    _write_csv(output_dir / "2025_emission_regression.csv", emission_rows)
    _write_csv(output_dir / "2025_qc_regression.csv", qc_rows)
    _write_csv(output_dir / "2025_lineage_regression.csv", lineage_rows)
    _write_csv(output_dir / "2025_scope_comparison.csv", scope_rows)
    _write_csv(output_dir / "2025_canonical_results.csv", canonical_rows)
    rounding = {
        "Unrounded_Total": current_summary["raw_total_emission_kgco2e"],
        "Rounded_Aggregate_Total": current_summary["official_six_decimal_total"],
        "Sum_Of_Row_Rounded": current_summary["row_six_decimal_sum"],
        "Rounding_Reconciliation": current_summary["rounding_reconciliation_difference"],
        "Rounding_Mode": "ROUND_HALF_UP",
        "Validation_Status": "PASS" if strict_checks["Rounding_Match"] else "FAIL",
    }
    _write_json(output_dir / "2025_rounding_reconciliation.json", rounding)
    _write_json(output_dir / "2025_enterprise_scope_reconciliation.json", scope_reconciliation)
    _write_json(output_dir / "2024_shared_model_forward_compatibility.json", wp63_forward)

    after_hashes = {name: _sha256(path) for name, path in sources.items()}
    hash_checks = {name: before_hashes[name] == after_hashes[name] for name in sources}
    status = "PASS" if strict_status == "STRICT_REGRESSION_PASS" and all(hash_checks.values()) else "BLOCKED"
    summary = {
        "schema_version": "WP6_4_SUMMARY_V1",
        "stage": "WP6-4",
        "status": status,
        "run_id": current_run_id,
        "current_run_id": current.name,
        "source_file": recognized["workbook_name"],
        "source_sha256": recognized["input_fingerprint"],
        "processing_policy_selected_path": "PCS_WEIGHT_DERIVED",
        "boundary": {
            "business_unit_count": boundary_audit[0]["Selected_Count"],
            "cell_count": boundary_audit[1]["Selected_Count"],
            "marker_count": boundary_audit[2]["Selected_Count"],
        },
        "strict_regression": {
            "status": strict_status,
            "record_count": len(current_ids),
            "checks": strict_checks,
            "activity_total_kg": current_summary["activity_total_kg"],
            "unrounded_total_kgco2e": current_summary["raw_total_emission_kgco2e"],
            "rounded_aggregate_total": current_summary["official_six_decimal_total"],
            "sum_of_row_rounded": current_summary["row_six_decimal_sum"],
            "rounding_reconciliation": current_summary["rounding_reconciliation_difference"],
        },
        "quality": quality_summary,
        "ef_audit": ef_summary,
        "enterprise_scope_reconciliation": scope_reconciliation,
        "wp6_3_forward_compatibility": wp63_forward,
        "protected_input_hashes_before": before_hashes,
        "protected_input_hashes_after": after_hashes,
        "protected_inputs_unchanged": hash_checks,
        "raw_data_modified": False,
        "frozen_evidence_modified": False,
        "wp6_5_execution_performed": False,
    }
    _write_json(output_dir / "wp6_4_summary.json", summary)
    (output_dir / "WP6-4_2025严格回归与范围勾稽报告.md").write_text(_build_report(summary), encoding="utf-8")
    return {
        "stage": "WP6-4",
        "status": status,
        "run_id": current_run_id,
        "output_directory": str(output_dir),
        "strict_regression_status": strict_status,
        "record_count": len(current_ids),
        "activity_coverage_percent": scope_reconciliation["activity_coverage_percent"],
        "emission_coverage_percent": scope_reconciliation["emission_coverage_percent"],
        "wp6_5_execution_performed": False,
    }
