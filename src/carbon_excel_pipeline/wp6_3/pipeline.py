"""File-backed WP6-3 pipeline consuming only WP6-1/2 structured artifacts."""

from __future__ import annotations

import csv
import json
from collections import Counter
from datetime import UTC, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.capability.models import ActivityPath
from carbon_excel_pipeline.capability.policy import select_activity_path
from carbon_excel_pipeline.errors import PipelineUserError

from .processing import (
    FORBIDDEN_FORWARD_FILL_FIELDS,
    NON_CRITICAL_ANALYSIS_FIELDS,
    adapt_direct_mass,
    adapt_historical_ef,
    adapt_historical_ghg,
    assign_2024_record_ids,
    calculate_and_validate,
    controlled_forward_fill,
    decimal_text,
    filter_2024_boundary,
    is_present,
)


OUTPUT_FIELDS = [
    "Record_ID",
    "Year",
    "Source_File",
    "Source_SHA256",
    "Source_Sheet",
    "Source_Row",
    "Business_Unit",
    "Purchase_Type",
    "Purchase_Category",
    "Product_Description",
    "Business_Unit_Original",
    "Purchase_Type_Original",
    "Purchase_Category_Original",
    "Business_Unit_Source",
    "Purchase_Type_Source",
    "Purchase_Category_Source",
    "Original_Activity_Value",
    "Original_Activity_Unit",
    "Activity_Data_kg",
    "Activity_Method",
    "Unit_Conversion_Factor",
    "EF_Value",
    "EF_Unit",
    "EF_Source",
    "EF_Usage",
    "Emission_kgCO2e",
    "Emission_tCO2e",
    "Historical_GHG_Value",
    "Historical_GHG_Unit",
    "Historical_Difference",
    "QC_Status",
    "Warning_Codes",
    "Blocking_Codes",
    "Simulation_Flag",
    "Production_Eligible",
    "Run_ID",
]


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def _write_json(path: Path, payload: Any) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _error(code: str, message: str, value: Any, suggestion: str) -> PipelineUserError:
    return PipelineUserError(
        stage="WP6_3_HISTORICAL_REPRODUCTION",
        error_code=code,
        message_cn=message,
        source_location="WP6-1/2结构化产物",
        original_value=value,
        rule="WP6-3只消费一致的Recognition、recognized_records和Capability结果，不重新扫描Excel。",
        impact="阻断WP6-3正式处理，但不改变上游、原始文件或冻结证据。",
        fix_suggestion=suggestion,
    )


def _validate_upstream(
    records: dict[str, Any],
    dataset: dict[str, Any],
    capabilities: list[dict[str, str]],
    decisions: list[dict[str, str]],
) -> dict[int, dict[str, str]]:
    source_rows = [int(item["Source_Row"]) for item in records.get("records", [])]
    capability_rows = [int(item["Source_Row"]) for item in capabilities]
    checks = {
        "fingerprints_match": records.get("input_fingerprint")
        == dataset.get("input_fingerprint"),
        "sheets_match": records.get("sheet_name") == dataset.get("sheet_name"),
        "headers_match": records.get("header_row") == dataset.get("header_row"),
        "record_count_matches": len(source_rows) == dataset.get("total_records"),
        "source_rows_unique": len(source_rows) == len(set(source_rows)),
        "capability_rows_match": set(source_rows) == set(capability_rows),
    }
    if not all(checks.values()):
        raise _error(
            "UPSTREAM_ARTIFACT_MISMATCH",
            "WP6-1/2结构化产物的指纹、工作表、行号或记录数不一致。",
            checks,
            "重新使用同一WP6-1 Run生成WP6-2产物后再执行。",
        )
    supported_pairs = {
        (int(item["Source_Row"]), item["Activity_Path"])
        for item in decisions
        if item.get("Supported", "").casefold() == "true"
    }
    capability_map = {int(item["Source_Row"]): item for item in capabilities}
    for row, item in capability_map.items():
        paths = [part for part in item.get("Supported_Activity_Paths", "").split("|") if part]
        for path in paths:
            if (row, path) not in supported_pairs:
                raise _error(
                    "CAPABILITY_PATH_DECISION_MISMATCH",
                    "记录能力表与路径判断表不一致。",
                    {"Source_Row": row, "Activity_Path": path},
                    "重新运行WP6-2能力识别。",
                )
    return capability_map


def _make_run_id(fingerprint: str) -> str:
    timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
    return f"WP6-3-{timestamp}-{fingerprint[:8]}"


def _baseline_comparison(
    actual: dict[str, Any],
    baseline: dict[str, Any],
    *,
    decimal_reporting_threshold: Decimal,
) -> dict[str, dict[str, Any]]:
    fields = {
        "business_unit_count": int,
        "cell_count": int,
        "marker_count": int,
        "original_activity_t": Decimal,
        "activity_kg": Decimal,
        "weighted_historical_ef": Decimal,
        "calculated_emission_tco2e": Decimal,
        "historical_emission_tco2e": Decimal,
    }
    output: dict[str, dict[str, Any]] = {}
    for field, converter in fields.items():
        actual_value = converter(actual[field])
        baseline_field = (
            "historical_emission_tco2e"
            if field == "calculated_emission_tco2e"
            else field
        )
        expected_value = converter(baseline[baseline_field])
        difference = actual_value - expected_value
        within_threshold = (
            difference == 0
            if converter is int
            else abs(difference) <= decimal_reporting_threshold
        )
        output[field] = {
            "actual": str(actual_value),
            "expected": str(expected_value),
            "difference": str(difference),
            "matches": difference == 0,
            "within_reporting_threshold": within_threshold,
        }
    return output


def _build_report(summary: dict[str, Any]) -> str:
    boundary = summary["boundary"]
    totals = summary["totals"]
    return "\n".join(
        [
            "# WP6-3 2024历史复现报告",
            "",
            f"- Run ID：`{summary['run_id']}`",
            f"- 输入文件：`{summary['source_file']}`",
            f"- 输入 SHA-256：`{summary['source_sha256']}`",
            f"- 上游 Run：`{summary['upstream_run_id']}`",
            f"- 状态：`{summary['status']}`",
            "",
            "## 业务边界",
            "",
            f"- 二部：{boundary['business_unit_count']} 条",
            f"- 电芯：{boundary['cell_count']} 条",
            f"- SYNA：{boundary['marker_count']} 条",
            "- 规则为字段精确/明确标记筛选，未使用固定行号。",
            "",
            "## 活动数据与核算",
            "",
            "- Selected Path：`DIRECT_REPORTED_MASS`",
            f"- 原始活动数据：{totals['original_activity_t']} t/year",
            f"- 标准活动数据：{totals['activity_kg']} kg/year",
            f"- 加权历史 EF：{totals['weighted_historical_ef']} kgCO2e/kg",
            f"- 重新计算排放：{totals['calculated_emission_tco2e']} tCO2e/year",
            f"- 历史排放：{totals['historical_emission_tco2e']} tCO2e/year",
            f"- 实际差异：{totals['difference_tco2e']} tCO2e/year",
            "",
            "## QC 与限制",
            "",
            f"- PASS：{summary['qc_counts'].get('PASS', 0)}",
            f"- WARNING：{summary['qc_counts'].get('WARNING', 0)}",
            f"- BLOCKED：{summary['qc_counts'].get('BLOCKED', 0)}",
            "- Chemistry、Supplier、Project、Model 缺失仅形成 Warning，未推断或填充。",
            "- Historical GHG 只参与事后比较，未进入 Activity × EF 计算。",
            "- Simulation_Flag=TRUE；Production_Eligible=FALSE。",
            "",
        ]
    )


def run_wp6_3_historical_reproduction(
    upstream_run_dir: Path,
    *,
    output_root: Path,
    policy_config_path: Path,
    run_id: str | None = None,
    output_dir: Path | None = None,
) -> dict[str, Any]:
    upstream = upstream_run_dir.expanduser().resolve()
    required = {
        "recognized_records": upstream / "01_import" / "recognized_records.json",
        "dataset_capabilities": upstream / "02_capability" / "dataset_capabilities.json",
        "record_capabilities": upstream / "02_capability" / "record_capabilities.csv",
        "activity_path_decisions": upstream / "02_capability" / "activity_path_decisions.csv",
    }
    missing = [name for name, path in required.items() if not path.is_file()]
    if missing:
        raise _error(
            "UPSTREAM_ARTIFACT_MISSING",
            "缺少WP6-1/2正式结构化输入。",
            missing,
            "提供已完成WP6-1与WP6-2的同一Run目录。",
        )
    policy = _load_json(policy_config_path.expanduser().resolve())
    records = _load_json(required["recognized_records"])
    dataset = _load_json(required["dataset_capabilities"])
    capabilities = _read_csv(required["record_capabilities"])
    decisions = _read_csv(required["activity_path_decisions"])
    capability_map = _validate_upstream(records, dataset, capabilities, decisions)

    requested = ActivityPath(policy["selected_activity_path"])
    if requested != ActivityPath.DIRECT_REPORTED_MASS:
        raise _error(
            "WP6_3_POLICY_INVALID",
            "WP6-3 2024正式策略必须显式选择DIRECT_REPORTED_MASS。",
            str(requested),
            "恢复受控WP6-3 2024 Processing Policy。",
        )
    fill_fields = tuple(policy["controlled_forward_fill_fields"])
    if set(fill_fields) & set(FORBIDDEN_FORWARD_FILL_FIELDS):
        raise _error(
            "FORWARD_FILL_POLICY_INVALID",
            "处理策略试图填充禁止继承的业务数值或描述字段。",
            fill_fields,
            "只允许Business_Unit、Purchase_Type、Purchase_Category受控继承。",
        )
    filled = controlled_forward_fill(records["records"], fields=fill_fields)
    boundary = policy["boundary"]
    selected, boundary_audit = filter_2024_boundary(
        filled,
        business_unit=boundary["business_unit_exact"],
        category_root=boundary["purchase_category_root_exact"],
        product_marker=boundary["product_description_marker"],
    )
    selected = assign_2024_record_ids(selected, year=int(policy["year"]))

    fingerprint = str(records["input_fingerprint"])
    current_run_id = run_id or _make_run_id(fingerprint)
    resolved_output_dir = (
        output_dir.expanduser().resolve()
        if output_dir is not None
        else output_root.expanduser().resolve() / current_run_id
    )
    if resolved_output_dir.exists() and any(resolved_output_dir.iterdir()):
        raise _error(
            "OUTPUT_RUN_ALREADY_EXISTS",
            "目标WP6-3 Run目录已存在且非空。",
            str(resolved_output_dir),
            "使用新的Run ID，保留既有证据不覆盖。",
        )
    resolved_output_dir.mkdir(parents=True, exist_ok=True)

    units = records.get("units", {})
    canonical_rows: list[dict[str, Any]] = []
    activity_rows: list[dict[str, Any]] = []
    ef_rows: list[dict[str, Any]] = []
    calculation_rows: list[dict[str, Any]] = []
    validation_rows: list[dict[str, Any]] = []
    qc_rows: list[dict[str, Any]] = []
    id_rows: list[dict[str, Any]] = []
    threshold = Decimal(policy["historical_difference_reporting_threshold_tco2e"])

    for record in selected:
        row_number = int(record["Source_Row"])
        values = record.get("values", {})
        cap = capability_map[row_number]
        supported = [
            item for item in cap.get("Supported_Activity_Paths", "").split("|") if item
        ]
        warnings: list[str] = []
        blocking: list[str] = []
        try:
            selected_path = select_activity_path(supported, requested_path=requested)
        except PipelineUserError:
            selected_path = None
            blocking.append("REQUESTED_ACTIVITY_PATH_NOT_SUPPORTED")
        activity = adapt_direct_mass(
            values.get("Reported_Activity_Value"),
            units.get("Reported_Activity_Value"),
        )
        factor = adapt_historical_ef(
            values.get("EF_Value"), units.get("EF_Value"), values.get("EF_Source")
        )
        historical = adapt_historical_ghg(
            values.get("Historical_GHG_Value"), units.get("Historical_GHG_Value")
        )
        warnings.extend(activity["warnings"])
        warnings.extend(factor["warnings"])
        warnings.extend(historical["warnings"])
        blocking.extend(activity["blocking"])
        blocking.extend(factor["blocking"])
        for field in NON_CRITICAL_ANALYSIS_FIELDS:
            if not is_present(values.get(field)):
                warnings.append(f"{field.upper()}_MISSING")

        calculation: dict[str, Any] | None = None
        if not blocking and selected_path is not None:
            calculation = calculate_and_validate(
                activity_kg=activity["activity_kg"],
                ef_value=factor["normalized_value"],
                historical_tco2e=historical["normalized_tco2e"],
                difference_threshold_tco2e=threshold,
            )
        warnings = sorted(set(warnings))
        blocking = sorted(set(blocking))
        qc_status = "BLOCKED" if blocking else ("WARNING" if warnings else "PASS")
        context_original = record.get("original_context_values", {})
        context_sources = record.get("context_value_sources", {})
        canonical = {
            "Record_ID": record["Record_ID"],
            "Year": str(policy["year"]),
            "Source_File": records["workbook_name"],
            "Source_SHA256": fingerprint,
            "Source_Sheet": records["sheet_name"],
            "Source_Row": row_number,
            "Business_Unit": values.get("Business_Unit"),
            "Purchase_Type": values.get("Purchase_Type"),
            "Purchase_Category": values.get("Purchase_Category"),
            "Product_Description": values.get("Product_Description"),
            "Business_Unit_Original": context_original.get("Business_Unit"),
            "Purchase_Type_Original": context_original.get("Purchase_Type"),
            "Purchase_Category_Original": context_original.get("Purchase_Category"),
            "Business_Unit_Source": context_sources.get("Business_Unit"),
            "Purchase_Type_Source": context_sources.get("Purchase_Type"),
            "Purchase_Category_Source": context_sources.get("Purchase_Category"),
            "Original_Activity_Value": values.get("Reported_Activity_Value"),
            "Original_Activity_Unit": units.get("Reported_Activity_Value"),
            "Activity_Data_kg": decimal_text(activity["activity_kg"]),
            "Activity_Method": str(selected_path or requested),
            "Unit_Conversion_Factor": decimal_text(activity["conversion_factor"]),
            "EF_Value": decimal_text(factor["normalized_value"]),
            "EF_Unit": "kgCO2e/kg" if factor["normalized_value"] is not None else units.get("EF_Value"),
            "EF_Source": values.get("EF_Source"),
            "EF_Usage": "HISTORICAL_REPRODUCTION",
            "Emission_kgCO2e": decimal_text(calculation["emission_kg"] if calculation else None),
            "Emission_tCO2e": decimal_text(calculation["emission_t"] if calculation else None),
            "Historical_GHG_Value": decimal_text(historical["normalized_tco2e"]),
            "Historical_GHG_Unit": "tCO2e/year" if historical["normalized_tco2e"] is not None else units.get("Historical_GHG_Value"),
            "Historical_Difference": decimal_text(calculation["difference_t"] if calculation else None),
            "QC_Status": qc_status,
            "Warning_Codes": "|".join(warnings),
            "Blocking_Codes": "|".join(blocking),
            "Simulation_Flag": "TRUE",
            "Production_Eligible": "FALSE",
            "Run_ID": current_run_id,
        }
        canonical_rows.append(canonical)
        id_rows.append(
            {
                "Record_ID": record["Record_ID"],
                "Source_SHA256": fingerprint,
                "Source_Sheet": records["sheet_name"],
                "Source_Row": row_number,
            }
        )
        activity_rows.append(
            {key: canonical[key] for key in (
                "Record_ID", "Source_Row", "Original_Activity_Value",
                "Original_Activity_Unit", "Activity_Data_kg", "Activity_Method",
                "Unit_Conversion_Factor", "QC_Status", "Warning_Codes", "Blocking_Codes"
            )}
        )
        ef_rows.append(
            {key: canonical[key] for key in (
                "Record_ID", "Source_Row", "EF_Value", "EF_Unit", "EF_Source",
                "EF_Usage", "Simulation_Flag", "Production_Eligible", "QC_Status"
            )}
        )
        qc_rows.append(
            {key: canonical[key] for key in (
                "Record_ID", "Source_Row", "QC_Status", "Warning_Codes", "Blocking_Codes"
            )}
        )
        if calculation:
            calculation_rows.append(
                {key: canonical[key] for key in (
                    "Record_ID", "Source_Row", "Activity_Data_kg", "EF_Value",
                    "Emission_kgCO2e", "Emission_tCO2e", "Run_ID"
                )}
            )
            validation_rows.append(
                {
                    "Record_ID": record["Record_ID"],
                    "Source_Row": row_number,
                    "Calculated_Emission_tCO2e": decimal_text(calculation["emission_t"]),
                    "Historical_Emission_tCO2e": decimal_text(historical["normalized_tco2e"]),
                    "Difference_tCO2e": decimal_text(calculation["difference_t"]),
                    "Difference_Percent": decimal_text(calculation["difference_percent"]),
                    "Validation_Status": calculation["validation_status"],
                }
            )

    calculated_activity = sum(
        (Decimal(row["Activity_Data_kg"]) for row in calculation_rows), Decimal("0")
    )
    calculated_emission = sum(
        (Decimal(row["Emission_tCO2e"]) for row in calculation_rows), Decimal("0")
    )
    historical_total = sum(
        (
            Decimal(row["Historical_Emission_tCO2e"])
            for row in validation_rows
            if row["Historical_Emission_tCO2e"]
        ),
        Decimal("0"),
    )
    original_activity_t = sum(
        (Decimal(str(row["Original_Activity_Value"])) for row in canonical_rows),
        Decimal("0"),
    )
    weighted_ef = (
        calculated_emission * Decimal("1000") / calculated_activity
        if calculated_activity
        else Decimal("0")
    )
    boundary_counts = {
        "business_unit_count": boundary_audit[0]["Selected_Count"],
        "cell_count": boundary_audit[1]["Selected_Count"],
        "marker_count": boundary_audit[2]["Selected_Count"],
    }
    totals = {
        "original_activity_t": decimal_text(original_activity_t),
        "activity_kg": decimal_text(calculated_activity),
        "weighted_historical_ef": decimal_text(weighted_ef),
        "calculated_emission_tco2e": decimal_text(calculated_emission),
        "historical_emission_tco2e": decimal_text(historical_total),
        "difference_tco2e": decimal_text(calculated_emission - historical_total),
    }
    baseline_actual = {
        **boundary_counts,
        "original_activity_t": totals["original_activity_t"],
        "activity_kg": totals["activity_kg"],
        "weighted_historical_ef": totals["weighted_historical_ef"],
        "calculated_emission_tco2e": totals["calculated_emission_tco2e"],
        "historical_emission_tco2e": totals["historical_emission_tco2e"],
    }
    baseline_checks = _baseline_comparison(
        baseline_actual,
        policy["regression_baseline"],
        decimal_reporting_threshold=threshold,
    )
    qc_counts = dict(Counter(row["QC_Status"] for row in canonical_rows))
    blocked_count = qc_counts.get("BLOCKED", 0)
    difference_fields = [
        field for field, check in baseline_checks.items() if not check["matches"]
    ]
    unexplained_fields = [
        field
        for field, check in baseline_checks.items()
        if not check["within_reporting_threshold"]
    ]
    status = "PASS_WITH_WARNING"
    if blocked_count or unexplained_fields:
        status = "PARTIAL_RESULT"
    summary = {
        "schema_version": "WP6_3_SUMMARY_V1",
        "stage": "WP6-3",
        "status": status,
        "run_id": current_run_id,
        "upstream_run_id": upstream.name,
        "upstream_run_directory": str(upstream),
        "source_file": records["workbook_name"],
        "source_sha256": fingerprint,
        "source_sheet": records["sheet_name"],
        "source_row_definition": "Original Excel worksheet physical row number, 1-based.",
        "capability_detector_role": "SUPPORTED_PATHS_ONLY",
        "processing_policy_selected_path": str(requested),
        "boundary": boundary_counts,
        "totals": totals,
        "qc_counts": qc_counts,
        "validation_status_counts": dict(
            Counter(row["Validation_Status"] for row in validation_rows)
        ),
        "baseline_checks": baseline_checks,
        "baseline_difference_fields": difference_fields,
        "baseline_unexplained_fields": unexplained_fields,
        "record_id_first": canonical_rows[0]["Record_ID"] if canonical_rows else None,
        "record_id_last": canonical_rows[-1]["Record_ID"] if canonical_rows else None,
        "calculated_record_count": len(calculation_rows),
        "historical_validation_count": len(validation_rows),
        "raw_excel_opened": False,
        "wp6_1_rerun": False,
        "wp6_2_rerun": False,
        "wp6_4_execution_performed": False,
    }

    _write_csv(
        resolved_output_dir / "2024_boundary_filter.csv",
        boundary_audit,
        ["Stage", "Input_Count", "Selected_Count", "Excluded_Count", "Filter_Rule"],
    )
    _write_csv(resolved_output_dir / "2024_record_id_mapping.csv", id_rows, list(id_rows[0]) if id_rows else ["Record_ID", "Source_SHA256", "Source_Sheet", "Source_Row"])
    _write_csv(resolved_output_dir / "2024_activity_data.csv", activity_rows, list(activity_rows[0]) if activity_rows else [])
    _write_csv(resolved_output_dir / "2024_qc_results.csv", qc_rows, list(qc_rows[0]) if qc_rows else [])
    _write_csv(resolved_output_dir / "2024_historical_ef.csv", ef_rows, list(ef_rows[0]) if ef_rows else [])
    _write_csv(resolved_output_dir / "2024_calculation_results.csv", calculation_rows, list(calculation_rows[0]) if calculation_rows else ["Record_ID", "Source_Row", "Activity_Data_kg", "EF_Value", "Emission_kgCO2e", "Emission_tCO2e", "Run_ID"])
    _write_csv(resolved_output_dir / "2024_historical_validation.csv", validation_rows, list(validation_rows[0]) if validation_rows else ["Record_ID", "Source_Row", "Calculated_Emission_tCO2e", "Historical_Emission_tCO2e", "Difference_tCO2e", "Difference_Percent", "Validation_Status"])
    _write_csv(resolved_output_dir / "2024_canonical_results.csv", canonical_rows, OUTPUT_FIELDS)
    _write_json(resolved_output_dir / "wp6_3_summary.json", summary)
    (resolved_output_dir / "WP6-3_2024历史复现报告.md").write_text(
        _build_report(summary), encoding="utf-8"
    )
    return {
        "stage": "WP6-3",
        "status": status,
        "run_id": current_run_id,
        "output_directory": str(resolved_output_dir),
        "record_count": len(canonical_rows),
        "calculated_record_count": len(calculation_rows),
        "baseline_difference_fields": difference_fields,
        "baseline_unexplained_fields": unexplained_fields,
        "wp6_4_execution_performed": False,
    }
