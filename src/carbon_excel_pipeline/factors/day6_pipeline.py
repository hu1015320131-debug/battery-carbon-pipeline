"""Orchestrate Day 6 historical factor receipt, D1 adaptation and exact matching."""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.errors import PipelineUserError
from carbon_excel_pipeline.factors.day6_adapter import (
    ADAPTATION_AUDIT_FIELDS,
    adapt_to_d1,
    candidate_anomalies,
    generate_historical_eight_field_input,
    load_json,
    receive_factor_input,
    sha256,
    write_csv,
)
from carbon_excel_pipeline.matching.day6_exact import ANOMALY_FIELDS, exact_match_and_route


def _error(code: str, message: str, location: str, value: Any, rule: str) -> PipelineUserError:
    return PipelineUserError(
        stage="DAY6_FACTOR_MATCHING",
        error_code=code,
        message_cn=message,
        source_location=location,
        original_value=value,
        rule=rule,
        impact="阻断Day 6正式精确匹配基线",
        fix_suggestion="恢复通过G1A的活动数据并按Day 6契约重新运行。",
    )


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return list(reader.fieldnames or []), list(reader)


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def run_day6_factors_and_matching(
    run_dir: Path,
    *,
    profile_config_path: Path,
    activity_contract_path: Path,
    external_contract_path: Path,
    d1_contract_path: Path,
    factor_config_path: Path,
    d2_contract_path: Path,
    route_contract_path: Path,
    factor_input_path: Path | None = None,
    historical_simulation: bool = False,
) -> dict[str, Any]:
    run_dir = run_dir.expanduser().resolve()
    activity_dir = run_dir / "05_activity"
    factor_dir = run_dir / "07_factor_results"
    matching_dir = run_dir / "08_matching"
    for directory in (activity_dir, factor_dir, matching_dir):
        if not directory.is_dir():
            raise _error(
                "DAY6_STAGE_DIRECTORY_MISSING", "运行目录缺少Day 6阶段目录。", str(directory), "MISSING", "Day 6必须接续受控创建的完整运行目录。"
            )
    gate_path = activity_dir / "day5_g1a_gate.json"
    activity_path = activity_dir / "day5_activity_36_fields.csv"
    if not gate_path.is_file() or not activity_path.is_file():
        raise _error(
            "DAY5_G1A_EVIDENCE_MISSING", "缺少Day 5活动数据或G1A证据。", str(activity_dir), "MISSING", "Day 6只接收已通过G1A的36字段活动数据。"
        )
    gate = load_json(gate_path)
    if gate.get("status") != "PASS" or gate.get("gate_status") != "G1A_UPSTREAM_REBUILD_RECONCILED":
        raise _error(
            "G1A_NOT_PASSED", "G1A门禁未通过。", gate_path.name, gate.get("gate_status"), "私有Day 6必须以G1A_UPSTREAM_REBUILD_RECONCILED为前置条件。"
        )
    profile = load_json(profile_config_path)
    if profile.get("classification") != "PUBLIC_SYNTHETIC_ONLY":
        raise _error(
            "DAY6_PROFILE_NOT_PUBLIC",
            "当前 Day 6 仅允许公开合成配置。",
            profile_config_path.name,
            profile.get("classification"),
            "公开仓不得加载私有业务 Profile。",
        )
    activity_contract = load_json(activity_contract_path)
    activity_fields, activity_records = _read_csv(activity_path)
    if activity_fields != activity_contract["fields"] or len(activity_fields) != 36:
        raise _error(
            "ACTIVITY_36_SCHEMA_MISMATCH", "活动数据不是冻结的36字段结构。", activity_path.name, len(activity_fields), "字段名称、数量和顺序必须与活动契约一致。"
        )
    if not activity_records:
        raise _error(
            "ACTIVITY_COUNT_INVALID", "Day 6 活动记录为空。", activity_path.name, 0, "至少需要一条合成活动记录。"
        )

    external_contract = load_json(external_contract_path)
    d1_contract = load_json(d1_contract_path)
    factor_config = load_json(factor_config_path)
    d2_contract = load_json(d2_contract_path)
    route_contract = load_json(route_contract_path)
    counts = {
        "external": len(external_contract["fields"]),
        "d1": len(d1_contract["fields"]),
        "d2": len(d2_contract["fields"]),
        "route": len(route_contract["fields"]),
    }
    if counts != {"external": 8, "d1": 45, "d2": 57, "route": 36}:
        raise _error(
            "DAY6_CONTRACT_FIELD_COUNT_INVALID", "Day 6字段契约数量不正确。", "config", counts, "字段数量必须严格为8/45/57/36。"
        )
    if historical_simulation == (factor_input_path is not None):
        raise _error(
            "FACTOR_SOURCE_MODE_INVALID", "必须且只能选择历史模拟或一个外部因子文件。", "Day6 arguments", {"historical_simulation": historical_simulation, "factor_input": str(factor_input_path or "")}, "两个输入模式互斥且必须选择一个。"
        )
    if historical_simulation:
        generated_path = factor_dir / "day6_historical_simulation_8_fields.csv"
        factor_input_path = generate_historical_eight_field_input(
            activity_records,
            config=factor_config,
            output_path=generated_path,
            fields=external_contract["fields"],
        )
    assert factor_input_path is not None
    input_hash_before = sha256(factor_input_path)
    _, factor_records, metadata = receive_factor_input(
        factor_input_path, stage_dir=factor_dir, contract=external_contract
    )
    anomalies = candidate_anomalies(activity_records, factor_records)
    write_csv(matching_dir / "day6_match_anomalies.csv", anomalies, ANOMALY_FIELDS)
    write_csv(matching_dir / "day6_manual_review.csv", anomalies, ANOMALY_FIELDS)
    if anomalies:
        summary = {
            "run_id": run_dir.name,
            "status": "FAIL",
            "stage_status": "DAY6_CANDIDATE_CARDINALITY_BLOCKED",
            "profile_id": profile["profile_id"],
            "activity_records": len(activity_records),
            "factor_input_records": len(factor_records),
            "anomaly_records": len(anomalies),
            "manual_review_records": len(anomalies),
            "fallback_attempted_records": 0,
            "receipt_sha256": metadata["receipt_sha256"],
        }
        _write_json(matching_dir / "day6_summary.json", summary)
        return {**summary, "run_directory": str(run_dir)}

    d1_records, audit_records = adapt_to_d1(
        factor_records,
        activity_records,
        metadata=metadata,
        factor_config=factor_config,
        target_fields=d1_contract["fields"],
    )
    d1_path = factor_dir / "day6_d1_factor_results_45_fields.csv"
    write_csv(d1_path, d1_records, d1_contract["fields"])
    write_csv(
        factor_dir / "day6_8_to_45_adaptation_audit.csv",
        audit_records,
        ADAPTATION_AUDIT_FIELDS,
    )
    d2_records, route_records, match_anomalies = exact_match_and_route(
        activity_records,
        d1_records,
        d2_fields=d2_contract["fields"],
        route_fields=route_contract["fields"],
        run_id=run_dir.name,
        evaluated_at=metadata["received_at"],
    )
    write_csv(matching_dir / "day6_d2_exact_matches_57_fields.csv", d2_records, d2_contract["fields"])
    write_csv(matching_dir / "day6_d3_exact_routes_36_fields.csv", route_records, route_contract["fields"])
    if match_anomalies:
        write_csv(matching_dir / "day6_match_anomalies.csv", match_anomalies, ANOMALY_FIELDS)
        write_csv(matching_dir / "day6_manual_review.csv", match_anomalies, ANOMALY_FIELDS)

    input_unchanged = sha256(factor_input_path) == input_hash_before
    exact_locked = sum(row["Exact_Lock_Status"] == "EXACT_LOCKED" for row in route_records)
    fallback_attempted = sum(row["Fallback_Attempted"] == "TRUE" for row in route_records)
    manual_review = len(match_anomalies)
    n = len(activity_records)
    status = "PASS" if (
        len(d1_records) == len(d2_records) == len(route_records) == n
        and exact_locked == n
        and fallback_attempted == 0
        and manual_review == 0
        and input_unchanged
    ) else "FAIL"
    summary = {
        "run_id": run_dir.name,
        "status": status,
        "stage_status": "DAY6_HISTORICAL_FACTOR_EXACT_MATCH_LOCKED" if status == "PASS" else "DAY6_BLOCKED",
        "profile_id": profile["profile_id"],
        "g1a_gate_status": gate["gate_status"],
        "activity_records": len(activity_records),
        "factor_input_records": len(factor_records),
        "factor_input_field_count": len(external_contract["fields"]),
        "d1_records": len(d1_records),
        "d1_field_count": len(d1_contract["fields"]),
        "d2_exact_match_records": len(d2_records),
        "d2_field_count": len(d2_contract["fields"]),
        "d3_route_records": len(route_records),
        "d3_field_count": len(route_contract["fields"]),
        "exact_locked_records": exact_locked,
        "fallback_attempted_records": fallback_attempted,
        "anomaly_records": len(match_anomalies),
        "manual_review_records": manual_review,
        "ef_value": factor_config["ef_value"],
        "ef_unit": factor_config["ef_unit"],
        "simulation_flag": "TRUE",
        "synthetic_test_flag": factor_config["synthetic_test_flag"],
        "production_eligible": factor_config["production_eligible"],
        "receipt_sha256": metadata["receipt_sha256"],
        "factor_input_unchanged": input_unchanged,
        "outputs": {
            "raw_receipt": metadata["raw_payload_path"],
            "d1_45_fields": "07_factor_results/day6_d1_factor_results_45_fields.csv",
            "adaptation_audit": "07_factor_results/day6_8_to_45_adaptation_audit.csv",
            "d2_57_fields": "08_matching/day6_d2_exact_matches_57_fields.csv",
            "d3_route_36_fields": "08_matching/day6_d3_exact_routes_36_fields.csv",
            "anomalies": "08_matching/day6_match_anomalies.csv",
            "manual_review": "08_matching/day6_manual_review.csv"
        },
    }
    _write_json(matching_dir / "day6_summary.json", summary)
    return {**summary, "run_directory": str(run_dir)}
