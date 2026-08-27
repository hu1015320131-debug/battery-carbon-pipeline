"""Orchestrate Day 5 quality, activity, third-party input and G1A."""

from __future__ import annotations

import csv
import hashlib
import json
from decimal import Decimal
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.activity.day5_builders import (
    RECORD_OPEN_ITEM_FIELDS,
    build_activity_records,
    build_record_open_items,
    build_third_party_records,
)
from carbon_excel_pipeline.activity.g1a_gate import evaluate_g1a
from carbon_excel_pipeline.errors import PipelineUserError
from carbon_excel_pipeline.qc.day5_quality import run_quality_checks


QUALITY_ISSUE_FIELDS = [
    "Record_ID",
    "Source_File",
    "Source_Sheet",
    "Source_Row",
    "Rule_ID",
    "Severity",
    "Issue_Code",
    "Missing_Required_Fields",
]
BLOCKED_FIELDS = ["Record_ID", "Source_Row", "QC_Status", "Issue_Code", "Block_Reason"]


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, payload: Any) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return list(reader.fieldnames or []), list(reader)


def _write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def _pipeline_error(
    *, code: str, message: str, location: str, value: Any, rule: str, suggestion: str
) -> PipelineUserError:
    return PipelineUserError(
        stage="DAY5_UPSTREAM_REBUILD",
        error_code=code,
        message_cn=message,
        source_location=location,
        original_value=value,
        rule=rule,
        impact="阻断Day 5和后续因子阶段",
        fix_suggestion=suggestion,
    )


def run_day5_upstream_rebuild(
    run_dir: Path,
    *,
    profile_config_path: Path,
    standard_contract_path: Path,
    activity_contract_path: Path,
    third_party_contract_path: Path,
    quality_config_path: Path,
    interface_open_items_path: Path,
    standard_baseline_path: Path | None = None,
    activity_baseline_path: Path | None = None,
    third_party_baseline_path: Path | None = None,
) -> dict[str, Any]:
    run_dir = run_dir.expanduser().resolve()
    stage_dirs = {
        "standard": run_dir / "03_standardized",
        "qc": run_dir / "04_qc",
        "activity": run_dir / "05_activity",
        "third_party": run_dir / "06_third_party_input",
    }
    for name, path in stage_dirs.items():
        if not path.is_dir():
            raise _pipeline_error(
                code="DAY5_STAGE_DIRECTORY_MISSING",
                message="运行目录缺少Day 5所需阶段目录。",
                location=name,
                value=path.name,
                rule="Day 5必须接续受控创建的完整运行目录。",
                suggestion="重新执行Day 2接收流程。",
            )
    standard_path = stage_dirs["standard"] / "day4_standard_31_fields.csv"
    day4_summary_path = stage_dirs["standard"] / "day4_standardization_summary.json"
    if not standard_path.is_file() or not day4_summary_path.is_file():
        raise _pipeline_error(
            code="DAY4_OUTPUT_MISSING",
            message="缺少Day 4标准数据或阶段摘要。",
            location="03_standardized",
            value="day4_standard_31_fields.csv",
            rule="Day 5只能处理通过Day 4的31字段标准数据。",
            suggestion="先执行Day 4 standardize命令。",
        )
    day4_summary = _load_json(day4_summary_path)
    if day4_summary.get("status") != "PASS":
        raise _pipeline_error(
            code="DAY4_NOT_PASSED",
            message="Day 4阶段状态不是PASS。",
            location="day4_standardization_summary.json",
            value=day4_summary.get("status"),
            rule="31字段标准化必须先通过。",
            suggestion="修复Day 4验证错误后重新运行。",
        )

    profile = _load_json(profile_config_path)
    standard_contract = _load_json(standard_contract_path)
    activity_contract = _load_json(activity_contract_path)
    third_party_contract = _load_json(third_party_contract_path)
    quality_config = _load_json(quality_config_path)
    interface_config = _load_json(interface_open_items_path)
    standard_fields = [item["name"] for item in standard_contract["fields"]]
    activity_fields = activity_contract["fields"]
    third_party_fields = third_party_contract["fields"]
    if len(standard_fields) != 31 or len(activity_fields) != 36 or len(third_party_fields) != 20:
        raise _pipeline_error(
            code="UPSTREAM_SCHEMA_FIELD_COUNT_INVALID",
            message="31/36/20字段契约数量不正确。",
            location="config",
            value={
                "standard": len(standard_fields),
                "activity": len(activity_fields),
                "third_party": len(third_party_fields),
            },
            rule="WP2/WP3字段数量与顺序必须冻结。",
            suggestion="恢复受控版本的字段契约。",
        )
    if activity_fields[:31] != standard_fields:
        raise _pipeline_error(
            code="ACTIVITY_SCHEMA_PREFIX_MISMATCH",
            message="36字段活动数据没有以31字段标准结构开头。",
            location=activity_contract_path.name,
            value="field order mismatch",
            rule="活动数据只能在31字段后追加5个派生字段。",
            suggestion="恢复WP3活动数据字段顺序。",
        )
    if profile["profile_id"] == "public_synthetic_profile" and not all(
        (standard_baseline_path, activity_baseline_path, third_party_baseline_path)
    ):
        raise _pipeline_error(
            code="PRIVATE_G1A_BASELINES_REQUIRED",
            message="私有正式G1A缺少WP2/WP3只读基线。",
            location="Day5 arguments",
            value="missing baseline path",
            rule="私有G1A必须在生成后与31/36/20冻结基线比较。",
            suggestion="提供三个--*-baseline参数。",
        )

    input_fields, standard_records = _read_csv(standard_path)
    if input_fields != standard_fields:
        raise _pipeline_error(
            code="DAY4_STANDARD_SCHEMA_MISMATCH",
            message="Day 4标准数据字段名称或顺序不一致。",
            location=standard_path.name,
            value={"actual_field_count": len(input_fields)},
            rule="Day 5输入必须严格为WP2的31字段顺序。",
            suggestion="重新执行Day 4并检查标准化契约。",
        )
    input_hash_before = _sha256(standard_path)
    checked, quality_issues, quality_summary = run_quality_checks(
        standard_records, config=quality_config
    )
    activity_records, blocked_records = build_activity_records(
        checked,
        activity_fields=activity_fields,
        semantic_zero_tolerance_g=Decimal(
            quality_config["registered_float_tail_tolerance_g"]
        ),
    )
    third_party_records = build_third_party_records(
        activity_records, contract=third_party_contract
    )
    record_open_items = build_record_open_items(activity_records)
    interface_open_items = interface_config["items"]

    _write_csv(
        stage_dirs["qc"] / "day5_checked_standard_31_fields.csv",
        checked,
        standard_fields,
    )
    _write_csv(
        stage_dirs["qc"] / "day5_quality_issue_records.csv",
        quality_issues,
        QUALITY_ISSUE_FIELDS,
    )
    _write_csv(
        stage_dirs["qc"] / "day5_blocked_records.csv",
        blocked_records,
        BLOCKED_FIELDS,
    )
    _write_json(stage_dirs["qc"] / "day5_quality_summary.json", quality_summary)
    _write_csv(
        stage_dirs["activity"] / "day5_activity_36_fields.csv",
        activity_records,
        activity_fields,
    )
    _write_csv(
        stage_dirs["activity"] / "day5_record_open_items.csv",
        record_open_items,
        RECORD_OPEN_ITEM_FIELDS,
    )
    _write_csv(
        stage_dirs["activity"] / "day5_interface_open_items.csv",
        interface_open_items,
        interface_config["fields"],
    )
    _write_csv(
        stage_dirs["third_party"] / "day5_third_party_20_fields.csv",
        third_party_records,
        third_party_fields,
    )

    if profile["profile_id"] == "public_synthetic_profile":
        gate = evaluate_g1a(
            standard_records=checked,
            standard_fields=standard_fields,
            activity_records=activity_records,
            activity_fields=activity_fields,
            third_party_records=third_party_records,
            third_party_fields=third_party_fields,
            quality_counts=quality_summary["status_counts"],
            record_open_item_count=len(record_open_items),
            interface_open_item_count=len(interface_open_items),
            standard_baseline_path=standard_baseline_path,
            activity_baseline_path=activity_baseline_path,
            third_party_baseline_path=third_party_baseline_path,
            quality_config=quality_config,
        )
    else:
        gate = {
            "gate_status": "G1A_NOT_APPLICABLE_TO_PUBLIC_SYNTHETIC_FIXTURE",
            "status": "NOT_EVALUATED",
            "checks": {},
            "baselines_used_as_calculation_input": False,
            "row_values_exported": False,
        }
    _write_json(stage_dirs["activity"] / "day5_g1a_gate.json", gate)

    input_hash_after = _sha256(standard_path)
    status = "PASS"
    if quality_summary["status_counts"]["ERROR"] or not activity_records:
        status = "FAIL"
    if profile["profile_id"] == "public_synthetic_profile" and gate["status"] != "PASS":
        status = "FAIL"
    summary = {
        "run_id": run_dir.name,
        "status": status,
        "profile_id": profile["profile_id"],
        "quality_config_id": quality_config["config_id"],
        "standard_input_records": len(standard_records),
        "quality_status_counts": quality_summary["status_counts"],
        "activity_records": len(activity_records),
        "activity_field_count": len(activity_fields),
        "third_party_records": len(third_party_records),
        "third_party_field_count": len(third_party_fields),
        "blocked_records": len(blocked_records),
        "record_open_items": len(record_open_items),
        "interface_open_items": len(interface_open_items),
        "g1a_gate_status": gate["gate_status"],
        "day4_input_sha256": input_hash_before,
        "day4_input_unchanged": input_hash_before == input_hash_after,
        "outputs": {
            "quality_summary": "04_qc/day5_quality_summary.json",
            "activity_data": "05_activity/day5_activity_36_fields.csv",
            "third_party_input": "06_third_party_input/day5_third_party_20_fields.csv",
            "g1a_gate": "05_activity/day5_g1a_gate.json",
        },
    }
    _write_json(stage_dirs["activity"] / "day5_upstream_summary.json", summary)
    return {**summary, "run_directory": str(run_dir)}
