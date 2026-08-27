"""Orchestrate Day 7 D4/D5 calculation, reconciliation and two-layer lineage."""

from __future__ import annotations

import csv
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.calculation.day7_calculator import (
    CALCULATION_AUDIT_FIELDS,
    EMISSION_SUMMARY_FIELDS,
    build_d5_results,
    build_emission_summary,
    calculate_d4,
)
from carbon_excel_pipeline.errors import PipelineUserError
from carbon_excel_pipeline.lineage.day7_lineage import build_two_layer_lineage


BLOCKED_FIELDS = ["Record_ID", "Stage", "Issue_Code", "Block_Reason"]


def _error(code: str, message: str, location: str, value: Any, rule: str) -> PipelineUserError:
    return PipelineUserError(
        stage="DAY7_END_TO_END_CALCULATION",
        error_code=code,
        message_cn=message,
        source_location=location,
        original_value=value,
        rule=rule,
        impact="阻断Day 7正式核算和血缘输出",
        fix_suggestion="恢复Day 6通过状态和冻结字段契约后重新执行。",
    )


def _load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return list(reader.fieldnames or []), list(reader)


def _write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def _read_checked(path: Path, expected_fields: list[str], label: str) -> list[dict[str, str]]:
    if not path.is_file():
        raise _error("DAY7_INPUT_MISSING", "Day 7输入文件不存在。", label, str(path), "所有Day 5/6受控输入必须存在。")
    fields, records = _read_csv(path)
    if fields != expected_fields:
        raise _error("DAY7_INPUT_SCHEMA_MISMATCH", "Day 7输入字段名称或顺序不一致。", label, len(fields), "输入必须严格符合冻结字段契约。")
    return records


def run_day7_calculation_and_lineage(
    run_dir: Path,
    *,
    profile_config_path: Path,
    calculation_rules_path: Path,
    activity_contract_path: Path,
    d1_contract_path: Path,
    d2_contract_path: Path,
    d3_contract_path: Path,
    d4_contract_path: Path,
    d5_contract_path: Path,
    frozen_lineage_contract_path: Path,
    extended_lineage_contract_path: Path,
) -> dict[str, Any]:
    run_dir = run_dir.expanduser().resolve()
    calculation_dir = run_dir / "09_calculation"
    output_dir = run_dir / "10_output"
    for directory in (calculation_dir, output_dir):
        if not directory.is_dir():
            raise _error("DAY7_STAGE_DIRECTORY_MISSING", "运行目录缺少Day 7阶段目录。", directory.name, str(directory), "Day 7必须接续完整隔离运行目录。")
    day6_summary_path = run_dir / "08_matching/day6_summary.json"
    if not day6_summary_path.is_file():
        raise _error("DAY6_SUMMARY_MISSING", "缺少Day 6阶段摘要。", "08_matching", day6_summary_path.name, "Day 7只能接续通过的Day 6。")
    day6_summary = _load_json(day6_summary_path)
    required_day6 = {
        "status": "PASS",
        "stage_status": "DAY6_HISTORICAL_FACTOR_EXACT_MATCH_LOCKED",
        "fallback_attempted_records": 0,
        "manual_review_records": 0,
    }
    day6_mismatches = {
        key: {"actual": day6_summary.get(key), "expected": value}
        for key, value in required_day6.items()
        if day6_summary.get(key) != value
    }
    if day6_summary.get("exact_locked_records") != day6_summary.get("activity_records"):
        day6_mismatches["exact_locked_records"] = {
            "actual": day6_summary.get("exact_locked_records"),
            "expected": day6_summary.get("activity_records"),
        }
    if day6_mismatches:
        raise _error("DAY6_GATE_NOT_PASSED", "Day 6未满足核算前置门禁。", day6_summary_path.name, day6_mismatches, "必须全部精确锁定、0回退、0人工复核。")

    profile = _load_json(profile_config_path)
    if profile.get("classification") != "PUBLIC_SYNTHETIC_ONLY":
        raise _error("DAY7_PROFILE_NOT_PUBLIC", "当前 Day 7 仅允许公开合成 Profile。", profile_config_path.name, profile.get("classification"), "公开仓不得加载私有业务 Profile。")
    rules = _load_json(calculation_rules_path)
    contracts = {
        "activity": _load_json(activity_contract_path),
        "d1": _load_json(d1_contract_path),
        "d2": _load_json(d2_contract_path),
        "d3": _load_json(d3_contract_path),
        "d4": _load_json(d4_contract_path),
        "d5": _load_json(d5_contract_path),
        "frozen": _load_json(frozen_lineage_contract_path),
        "extended": _load_json(extended_lineage_contract_path),
    }
    expected_counts = {"activity": 36, "d1": 45, "d2": 57, "d3": 36, "d4": 48, "d5": 56, "frozen": 32, "extended": 25}
    actual_counts = {key: len(value["fields"]) for key, value in contracts.items()}
    if actual_counts != expected_counts:
        raise _error("DAY7_CONTRACT_FIELD_COUNT_INVALID", "Day 7契约字段数不正确。", "config", actual_counts, "字段数量必须为36/45/57/36/48/56/32/25。")
    if contracts["d5"]["fields"][8:] != contracts["d4"]["fields"]:
        raise _error("D5_D4_SCHEMA_RELATION_INVALID", "D5 56字段没有以8个端到端字段加D4 48字段组成。", d5_contract_path.name, "field order mismatch", "D5后48字段必须与D4完全一致。")

    paths = {
        "activity": run_dir / "05_activity/day5_activity_36_fields.csv",
        "third_party": run_dir / "06_third_party_input/day5_third_party_20_fields.csv",
        "d1": run_dir / "07_factor_results/day6_d1_factor_results_45_fields.csv",
        "d2": run_dir / "08_matching/day6_d2_exact_matches_57_fields.csv",
        "d3": run_dir / "08_matching/day6_d3_exact_routes_36_fields.csv",
    }
    activity = _read_checked(paths["activity"], contracts["activity"]["fields"], "activity")
    d1 = _read_checked(paths["d1"], contracts["d1"]["fields"], "d1")
    d2 = _read_checked(paths["d2"], contracts["d2"]["fields"], "d2")
    d3 = _read_checked(paths["d3"], contracts["d3"]["fields"], "d3")
    third_fields, third_party = _read_csv(paths["third_party"])
    if len(third_fields) != 20:
        raise _error("THIRD_PARTY_SCHEMA_INVALID", "第三方输入不是20字段。", paths["third_party"].name, len(third_fields), "Day 7血缘必须连接Day 5的20字段输入。")
    record_n = len(activity)
    if any(len(records) != record_n for records in (activity, d1, d2, d3, third_party)):
        raise _error("DAY7_INPUT_RECORD_COUNT_INVALID", "Day 7 各阶段记录数不一致。", "run inputs", {"activity": len(activity), "d1": len(d1), "d2": len(d2), "d3": len(d3), "third_party": len(third_party)}, "每个阶段的记录数必须一致。")

    input_hashes_before = {key: _sha256(path) for key, path in paths.items()}
    timestamp = datetime.now(timezone.utc).isoformat()
    d4, audit = calculate_d4(
        d3,
        rules=rules,
        d4_fields=contracts["d4"]["fields"],
        run_id=run_dir.name,
        calculated_at=timestamp,
    )
    emission_summary, reconciliation_checks = build_emission_summary(d4, rules=rules)
    d5 = build_d5_results(
        d4,
        d5_fields=contracts["d5"]["fields"],
        run_id=run_dir.name,
        executed_at=timestamp,
    )
    receipt = _load_json(run_dir / "01_import/file_receipt_report.json")
    metadata = {
        "activity_source_file": paths["activity"].name,
        "activity_source_sha256": input_hashes_before["activity"],
        "third_party_source_file": paths["third_party"].name,
        "third_party_source_sha256": input_hashes_before["third_party"],
        "raw_input_source_file": receipt["source_file_name"],
        "raw_input_sha256": receipt["source_sha256"],
        "received_input_copy": f"00_input_copy/{receipt['copy_file_name']}",
        "received_input_sha256": receipt["copy_sha256"],
        "import_run_id": run_dir.name,
        "profile_id": profile["profile_id"],
        "profile_config_sha256": _sha256(profile_config_path),
        "calculation_config_sha256": _sha256(calculation_rules_path),
    }
    frozen, extended, lineage_qa = build_two_layer_lineage(
        activity_records=activity,
        d1_records=d1,
        d2_records=d2,
        d3_records=d3,
        d4_records=d4,
        d5_records=d5,
        frozen_fields=contracts["frozen"]["fields"],
        extended_fields=contracts["extended"]["fields"],
        metadata=metadata,
    )

    d4_path = calculation_dir / "day7_d4_calculation_48_fields.csv"
    d5_path = output_dir / "day7_d5_end_to_end_56_fields.csv"
    _write_csv(d4_path, d4, contracts["d4"]["fields"])
    _write_csv(calculation_dir / "day7_calculation_audit_25_fields.csv", audit, CALCULATION_AUDIT_FIELDS)
    _write_csv(calculation_dir / "day7_emission_summary_16_fields.csv", [emission_summary], EMISSION_SUMMARY_FIELDS)
    _write_csv(calculation_dir / "day7_formal_blocked_records.csv", [], BLOCKED_FIELDS)
    _write_csv(d5_path, d5, contracts["d5"]["fields"])
    _write_csv(output_dir / "day7_wp5_frozen_lineage_32_fields.csv", frozen, contracts["frozen"]["fields"])
    _write_csv(output_dir / "day7_demo_extended_lineage_25_fields.csv", extended, contracts["extended"]["fields"])
    _write_json(output_dir / "day7_lineage_qa.json", lineage_qa)

    input_hashes_after = {key: _sha256(path) for key, path in paths.items()}
    input_unchanged = input_hashes_before == input_hashes_after
    status = "PASS" if (
        all(reconciliation_checks.values())
        and lineage_qa["status"] == "PASS"
        and len(d4) == len(d5) == len(frozen) == len(extended) == record_n
        and input_unchanged
    ) else "FAIL"
    summary = {
        "run_id": run_dir.name,
        "status": status,
        "stage_status": "DAY7_CALCULATION_LINEAGE_RECONCILED" if status == "PASS" else "DAY7_BLOCKED",
        "profile_id": profile["profile_id"],
        "day6_gate_status": day6_summary["stage_status"],
        "d4_records": len(d4),
        "d4_field_count": len(contracts["d4"]["fields"]),
        "d5_records": len(d5),
        "d5_field_count": len(contracts["d5"]["fields"]),
        "frozen_lineage_records": len(frozen),
        "frozen_lineage_field_count": len(contracts["frozen"]["fields"]),
        "extended_lineage_records": len(extended),
        "extended_lineage_field_count": len(contracts["extended"]["fields"]),
        "formal_blocked_records": 0,
        "activity_total_kg": emission_summary["Activity_Total_kg"],
        "ef_value": emission_summary["EF_Value"],
        "raw_total_emission_kgco2e": emission_summary["Raw_Total_Emission_kgCO2e"],
        "official_six_decimal_total": emission_summary["Total_Emission_kgCO2e"],
        "row_six_decimal_sum": emission_summary["Sum_of_Row_Display_Emission"],
        "rounding_reconciliation_difference": emission_summary["Rounding_Reconciliation_Difference"],
        "emission_unit": emission_summary["Emission_Unit"],
        "reconciliation_checks": reconciliation_checks,
        "lineage_qa_status": lineage_qa["status"],
        "input_files_unchanged": input_unchanged,
        "outputs": {
            "d4_48_fields": "09_calculation/day7_d4_calculation_48_fields.csv",
            "calculation_audit": "09_calculation/day7_calculation_audit_25_fields.csv",
            "emission_summary": "09_calculation/day7_emission_summary_16_fields.csv",
            "blocked_records": "09_calculation/day7_formal_blocked_records.csv",
            "d5_56_fields": "10_output/day7_d5_end_to_end_56_fields.csv",
            "frozen_lineage_32_fields": "10_output/day7_wp5_frozen_lineage_32_fields.csv",
            "extended_lineage": "10_output/day7_demo_extended_lineage_25_fields.csv",
            "lineage_qa": "10_output/day7_lineage_qa.json"
        },
    }
    _write_json(output_dir / "day7_summary.json", summary)
    return {**summary, "run_directory": str(run_dir)}
