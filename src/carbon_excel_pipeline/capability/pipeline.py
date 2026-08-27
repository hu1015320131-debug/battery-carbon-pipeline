"""File-backed WP6-2 detector consuming only WP6-1 structured artifacts."""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.errors import PipelineUserError

from .detector import detect_dataset_capabilities
from .models import CapabilityStatus


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, payload: Any) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _error(code: str, message: str, value: Any) -> PipelineUserError:
    return PipelineUserError(
        stage="CAPABILITY_DETECTION",
        error_code=code,
        message_cn=message,
        source_location="WP6-1结构化输出",
        original_value=value,
        rule="WP6-2只消费彼此一致的WP6-1识别与记录物化产物，不重新打开Excel。",
        impact="阻断本次能力判断，但不改变任何上游文件。",
        fix_suggestion="重新运行WP6-1结构识别以生成同一Run内的一致结构化产物。",
    )


def _validate_upstream(
    recognition: dict[str, Any],
    mappings: list[dict[str, Any]],
    records: dict[str, Any],
) -> None:
    best_sheet = recognition.get("best_candidate_sheet")
    best_header = recognition.get("best_candidate_header_row")
    mapped_sheet = next(
        (item for item in mappings if item.get("sheet_name") == best_sheet), None
    )
    checks = {
        "recognition_status": recognition.get("recognition_status"),
        "best_sheet": best_sheet,
        "best_header": best_header,
        "mapping_sheet_present": mapped_sheet is not None,
        "records_sheet": records.get("sheet_name"),
        "records_header": records.get("header_row"),
        "fingerprints_match": (
            recognition.get("input_fingerprint") == records.get("input_fingerprint")
        ),
    }
    if recognition.get("recognition_status") not in {
        "RECOGNIZED",
        "RECOGNIZED_WITH_WARNING",
    }:
        raise _error("UPSTREAM_RECOGNITION_NOT_READY", "上游尚未确认可用业务结构。", checks)
    if (
        mapped_sheet is None
        or records.get("sheet_name") != best_sheet
        or records.get("header_row") != best_header
        or not checks["fingerprints_match"]
    ):
        raise _error("UPSTREAM_ARTIFACT_MISMATCH", "上游结构化产物彼此不一致。", checks)
    mapped_units = {
        item["semantic_field"]: item.get("detected_unit")
        for item in mapped_sheet.get("field_mappings", [])
        if item.get("semantic_field")
    }
    if mapped_units != records.get("units", {}):
        raise _error(
            "UPSTREAM_UNIT_METADATA_MISMATCH",
            "语义映射与记录物化产物的单位元数据不一致。",
            {"mapping_units": mapped_units, "record_units": records.get("units", {})},
        )


def _record_csv_rows(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for item in records:
        statuses = item["Value_Statuses"]
        rows.append(
            {
                "Source_Sheet": item.get("Source_Sheet", ""),
                "Source_Row": item["Source_Row"],
                "Status": item["Status"],
                "Supported_Activity_Paths": "|".join(item["Supported_Activity_Paths"]),
                "Activity_Ready": item["Activity_Ready"],
                "Factor_Ready": item["Factor_Ready"],
                "Factor_Source_Available": item["Factor_Source_Available"],
                "Emission_Ready": item["Emission_Ready"],
                "Historical_Result_Available": item["Historical_Result_Available"],
                "Historical_Validation_Ready": item["Historical_Validation_Ready"],
                "PCS_Value_Status": statuses.get("Quantity_PCS", "MISSING"),
                "Unit_Weight_Value_Status": statuses.get("Unit_Weight", "MISSING"),
                "Reported_Activity_Value_Status": statuses.get(
                    "Reported_Activity_Value", "MISSING"
                ),
                "EF_Value_Status": statuses.get("EF_Value", "MISSING"),
                "Historical_GHG_Value_Status": statuses.get(
                    "Historical_GHG_Value", "MISSING"
                ),
                "Warning_Codes": "|".join(item["Warning_Codes"]),
                "Blocking_Codes": "|".join(item["Blocking_Codes"]),
                "Formula_Fields": "|".join(item["Formula_Fields"]),
            }
        )
    return rows


def _path_csv_rows(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for item in records:
        for decision in item["Path_Decisions"]:
            rows.append(
                {
                    "Source_Sheet": item.get("Source_Sheet", ""),
                    "Source_Row": item["Source_Row"],
                    "Activity_Path": decision["path"],
                    "Supported": decision["supported"],
                    "Reason_Codes": "|".join(decision["reason_codes"]),
                }
            )
    return rows


def _merge_cell_sheet_capabilities(
    primary_payload: dict[str, Any], extra_payloads: list[dict[str, Any]]
) -> dict[str, Any]:
    detected_sets = [
        detect_dataset_capabilities(payload)
        for payload in [primary_payload, *extra_payloads]
        if payload.get("records")
    ]
    if len(detected_sets) <= 1:
        return detected_sets[0] if detected_sets else detect_dataset_capabilities(primary_payload)
    count_keys = (
        "activity_ready_count",
        "factor_ready_count",
        "emission_ready_count",
        "historical_validation_ready_count",
        "pcs_weight_derived_count",
        "direct_reported_mass_count",
        "warning_count",
        "incapable_count",
    )
    total = sum(item["dataset"]["total_records"] for item in detected_sets)
    counts = {
        key: sum(int(item["dataset"].get(key) or 0) for item in detected_sets)
        for key in count_keys
    }
    if total == 0 or counts["activity_ready_count"] == 0:
        status = CapabilityStatus.INCAPABLE
    elif counts["emission_ready_count"] < total:
        status = CapabilityStatus.PARTIALLY_CAPABLE
    elif counts["warning_count"]:
        status = CapabilityStatus.CAPABLE_WITH_WARNING
    else:
        status = CapabilityStatus.CAPABLE
    analysis_keys = set().union(
        *(set(item["dataset"].get("analysis_capability_counts") or {}) for item in detected_sets)
    )
    analysis_counts = {
        key: sum(int(item["dataset"].get("analysis_capability_counts", {}).get(key) or 0) for item in detected_sets)
        for key in analysis_keys
    }
    dataset = {
        **detected_sets[0]["dataset"],
        "status": status,
        "sheet_name": "MULTI_SHEET_CELL_SCOPE",
        "header_row": None,
        "denominator_definition": "All recognized records inside the controlled cell boundary across activity sheets.",
        "total_records": total,
        **counts,
        **{
            key.removesuffix("_count") + "_coverage": value / total if total else 0.0
            for key, value in counts.items()
            if key.endswith("_count") and key not in {"warning_count", "incapable_count"}
        },
        "analysis_capability_counts": analysis_counts,
        "analysis_capability_coverage": {
            key: value / total if total else 0.0 for key, value in analysis_counts.items()
        },
        "sheets": [item["dataset"]["sheet_name"] for item in detected_sets],
    }
    records: list[dict[str, Any]] = []
    for detected in detected_sets:
        sheet_name = detected["dataset"].get("sheet_name")
        records.extend({"Source_Sheet": sheet_name, **row} for row in detected["records"])
    return {"dataset": dataset, "records": records}


def _markdown(dataset: dict[str, Any]) -> str:
    total = dataset["total_records"]
    return "\n".join(
        [
            "# WP6-2 数据能力识别报告",
            "",
            f"- 文件：`{dataset['workbook_name']}`",
            f"- 输入 SHA-256：`{dataset['input_fingerprint']}`",
            f"- 工作表：`{dataset['sheet_name']}`",
            f"- Header 行：{dataset['header_row']}",
            f"- 数据集能力状态：`{dataset['status']}`",
            f"- 覆盖率分母：{total} 条（{dataset['denominator_definition']}）",
            "",
            "## 准备度与路径覆盖",
            "",
            "|指标|记录数|覆盖率|",
            "|---|---:|---:|",
            f"|Activity Ready|{dataset['activity_ready_count']}|{dataset['activity_ready_coverage']:.2%}|",
            f"|Factor Ready|{dataset['factor_ready_count']}|{dataset['factor_ready_coverage']:.2%}|",
            f"|Emission Ready|{dataset['emission_ready_count']}|{dataset['emission_ready_coverage']:.2%}|",
            f"|Historical Validation Ready|{dataset['historical_validation_ready_count']}|{dataset['historical_validation_ready_coverage']:.2%}|",
            f"|PCS_WEIGHT_DERIVED|{dataset['pcs_weight_derived_count']}|{dataset['pcs_weight_derived_coverage']:.2%}|",
            f"|DIRECT_REPORTED_MASS|{dataset['direct_reported_mass_count']}|{dataset['direct_reported_mass_coverage']:.2%}|",
            "",
            "## 处理说明",
            "",
            "- 本报告仅判断能力与处理路径，不执行 Activity、Record_ID、EF Adapter 或 GHG 计算。",
            "- Detector 只输出 Supported Paths；Selected Path 由下游显式 Processing Policy 决定。",
            "- 单一路径失败不会阻止检查另一条路径；非关键分析字段缺失只形成 Warning。",
            "- WP6-2 只读取 WP6-1 结构化产物，没有重新打开或扫描原始 Excel。",
            "",
        ]
    )


def run_wp6_2_capability_detection(run_dir: Path) -> dict[str, Any]:
    run = run_dir.expanduser().resolve()
    import_dir = run / "01_import"
    output_dir = run / "02_capability"
    required = {
        "recognition_summary": import_dir / "recognition_summary.json",
        "semantic_field_mapping": import_dir / "semantic_field_mapping.json",
        "recognized_records": import_dir / "recognized_records.json",
    }
    missing = [name for name, path in required.items() if not path.is_file()]
    if missing:
        raise _error("UPSTREAM_ARTIFACT_MISSING", "缺少WP6-1结构化输入。", missing)
    recognition = _load_json(required["recognition_summary"])
    mappings = _load_json(required["semantic_field_mapping"])
    recognized_records = _load_json(required["recognized_records"])
    _validate_upstream(recognition, mappings, recognized_records)

    cell_scope_path = import_dir / "cell_scope_summary.json"
    extra_path = import_dir / "recognized_records_by_sheet.json"
    primary_fields = {
        item.get("semantic_field")
        for item in recognized_records.get("column_mappings") or []
    }
    multi_sheet_pcs_scope = {"Quantity_PCS", "Unit_Weight"}.issubset(primary_fields)
    if cell_scope_path.is_file() and extra_path.is_file() and multi_sheet_pcs_scope:
        extra_payload = _load_json(extra_path)
        detected = _merge_cell_sheet_capabilities(
            recognized_records,
            extra_payload.get("sheets") or [] if isinstance(extra_payload, dict) else extra_payload,
        )
    else:
        detected = detect_dataset_capabilities(recognized_records)
    dataset = detected["dataset"]
    records = detected["records"]
    output_dir.mkdir(parents=True, exist_ok=True)
    record_rows = _record_csv_rows(records)
    path_rows = _path_csv_rows(records)
    _write_json(output_dir / "dataset_capabilities.json", dataset)
    _write_json(
        output_dir / "capability_summary.json",
        {
            "stage": "WP6-2",
            "status": "PASS",
            "capability_status": dataset["status"],
            **{
                key: value
                for key, value in dataset.items()
                if key == "total_records"
                or key.endswith("_count")
                or key.endswith("_coverage")
            },
            "denominator_definition": dataset["denominator_definition"],
            "wp6_3_execution_performed": False,
        },
    )
    record_fields = list(record_rows[0]) if record_rows else [
        "Source_Sheet",
        "Source_Row",
        "Status",
        "Supported_Activity_Paths",
        "Activity_Ready",
        "Factor_Ready",
        "Factor_Source_Available",
        "Emission_Ready",
        "Historical_Result_Available",
        "Historical_Validation_Ready",
        "PCS_Value_Status",
        "Unit_Weight_Value_Status",
        "Reported_Activity_Value_Status",
        "EF_Value_Status",
        "Historical_GHG_Value_Status",
        "Warning_Codes",
        "Blocking_Codes",
        "Formula_Fields",
    ]
    _write_csv(output_dir / "record_capabilities.csv", record_rows, record_fields)
    _write_csv(
        output_dir / "activity_path_decisions.csv",
        path_rows,
        ["Source_Sheet", "Source_Row", "Activity_Path", "Supported", "Reason_Codes"],
    )
    (output_dir / "WP6-2_数据能力识别报告.md").write_text(
        _markdown(dataset), encoding="utf-8"
    )
    return {
        "stage": "WP6-2",
        "status": "PASS",
        "capability_status": dataset["status"],
        "run_id": run.name,
        "run_directory": str(run),
        "output_directory": str(output_dir),
        "total_records": dataset["total_records"],
        "activity_ready_count": dataset["activity_ready_count"],
        "emission_ready_count": dataset["emission_ready_count"],
        "wp6_3_execution_performed": False,
    }
