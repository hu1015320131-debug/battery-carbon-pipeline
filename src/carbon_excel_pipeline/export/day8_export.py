"""Day 8 curated CSV/JSON export and artifact-tool workbook orchestration."""

from __future__ import annotations

import csv
import hashlib
import json
import os
import shutil
import subprocess
from datetime import datetime, timezone
from pathlib import Path
from time import monotonic
from typing import Any

from carbon_excel_pipeline.errors import PipelineUserError


STATUS_OPEN_ITEM_FIELDS = [
    "Layer", "Item_ID", "Record_ID", "Status", "Issue_Code", "Owner",
    "Impact", "Evidence",
]
FINGERPRINT_FIELDS = [
    "Artifact_ID", "Artifact_Type", "File_Name", "Relative_Locator", "Size_Bytes",
    "SHA256", "Status",
]


def _error(code: str, message: str, location: str, value: Any, rule: str) -> PipelineUserError:
    return PipelineUserError(
        stage="DAY8_WORKBOOK_EXPORT",
        error_code=code,
        message_cn=message,
        source_location=location,
        original_value=value,
        rule=rule,
        impact="阻断Day 8工作簿、回读验证和G2门禁",
        fix_suggestion="恢复Day 7通过状态、artifact-tool运行环境和可写输出目录后重试。",
    )


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


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


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def _copy_required(source: Path, destination: Path) -> Path:
    if not source.is_file():
        raise _error("DAY8_SOURCE_MISSING", "Day 8导出源文件不存在。", source.name, str(source), "工作簿只能使用当前运行目录的完整阶段输出。")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    return destination


def _flatten_json(prefix: str, value: Any, rows: list[dict[str, str]]) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            _flatten_json(f"{prefix}.{key}" if prefix else str(key), child, rows)
    elif isinstance(value, list):
        rows.append({"Property": prefix, "Value": json.dumps(value, ensure_ascii=False), "Status": "RECORDED", "Evidence": "JSON_ARRAY"})
    else:
        rows.append({"Property": prefix, "Value": "" if value is None else str(value), "Status": "RECORDED", "Evidence": "JSON_VALUE"})


def _build_status_open_items(
    record_path: Path, interface_path: Path, governance_path: Path
) -> list[dict[str, str]]:
    _, record_rows = _read_csv(record_path)
    _, interface_rows = _read_csv(interface_path)
    _, governance_rows = _read_csv(governance_path)
    output: list[dict[str, str]] = []
    for row in record_rows:
        output.append({
            "Layer": "WP3_RECORD", "Item_ID": row["Open_Item_ID"],
            "Record_ID": row["Record_ID"], "Status": row["Status"],
            "Issue_Code": row["Source_Issue_Code"], "Owner": row["Owner"],
            "Impact": row["Impact"], "Evidence": row["Evidence"],
        })
    for row in interface_rows:
        output.append({
            "Layer": "WP3_INTERFACE", "Item_ID": row["Open_Item_ID"],
            "Record_ID": "", "Status": row["Current_Status"], "Issue_Code": row["Open_Item"],
            "Owner": row["Owner"], "Impact": row["Risk"], "Evidence": "DAY5_INTERFACE_REGISTER",
        })
    for row in governance_rows:
        output.append({
            "Layer": "WP5_GOVERNANCE", "Item_ID": row["Open_Item_ID"],
            "Record_ID": "", "Status": row["Status"], "Issue_Code": row["Topic"],
            "Owner": row["Owner"], "Impact": row["Impact"], "Evidence": f"WP5-D5:{row['Source']}",
        })
    return output


def _ensure_node_modules_junction(work_dir: Path, node_modules: Path) -> None:
    link = work_dir / "node_modules"
    if link.exists():
        if not (link / "@oai" / "artifact-tool").is_dir():
            raise _error("ARTIFACT_TOOL_JUNCTION_INVALID", "工作目录node_modules不包含artifact-tool。", str(link), "INVALID", "必须指向工作区加载器提供的node_modules。")
        return
    if os.name == "nt":
        completed = subprocess.run(
            ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(node_modules)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", check=False,
        )
        if completed.returncode != 0:
            raise _error("ARTIFACT_TOOL_JUNCTION_FAILED", "无法创建artifact-tool目录联接。", str(link), completed.stderr or completed.stdout, "Day 8必须在独立工作目录使用加载器提供的node_modules。")
    else:
        link.symlink_to(node_modules, target_is_directory=True)
    if not (link / "@oai" / "artifact-tool").is_dir():
        raise _error("ARTIFACT_TOOL_UNAVAILABLE", "artifact-tool不可用。", str(link), "MISSING", "不得安装替代库或使用非加载器依赖。")


def _fingerprints(paths: list[tuple[str, str, Path]], root: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for artifact_id, artifact_type, path in paths:
        rows.append({
            "Artifact_ID": artifact_id,
            "Artifact_Type": artifact_type,
            "File_Name": path.name,
            "Relative_Locator": str(path.relative_to(root)).replace("\\", "/") if path.is_relative_to(root) else path.name,
            "Size_Bytes": path.stat().st_size,
            "SHA256": _sha256(path),
            "Status": "RECORDED",
        })
    return rows


def run_day8_export(
    run_dir: Path,
    *,
    output_dir: Path,
    artifact_work_dir: Path,
    node_executable: Path,
    node_modules_path: Path,
    builder_script_path: Path,
    wp5_open_items_path: Path,
) -> dict[str, Any]:
    started = datetime.now(timezone.utc)
    timer = monotonic()
    run_dir = run_dir.expanduser().resolve()
    output_dir = output_dir.expanduser().resolve()
    artifact_work_dir = artifact_work_dir.expanduser().resolve()
    day7_summary_path = run_dir / "10_output/day7_summary.json"
    if not day7_summary_path.is_file():
        raise _error("DAY7_SUMMARY_MISSING", "缺少Day 7阶段摘要。", "10_output", day7_summary_path.name, "Day 8只能接续通过的Day 7。")
    day7 = _load_json(day7_summary_path)
    if day7.get("status") != "PASS" or day7.get("stage_status") != "DAY7_CALCULATION_LINEAGE_RECONCILED":
        raise _error("DAY7_GATE_NOT_PASSED", "Day 7未通过。", day7_summary_path.name, day7.get("stage_status"), "必须先完成核算、两层血缘和总量勾稽。")
    if not node_executable.is_file() or not node_modules_path.is_dir():
        raise _error("NODE_RUNTIME_MISSING", "加载器提供的Node运行环境不存在。", "Day8 arguments", {"node": str(node_executable), "modules": str(node_modules_path)}, "必须使用工作区依赖加载器返回的Node和node_modules。")
    if not builder_script_path.is_file() or not wp5_open_items_path.is_file():
        raise _error("DAY8_SUPPORT_FILE_MISSING", "Day 8构建器或WP5 Open Items不存在。", "Day8 arguments", {"builder": str(builder_script_path), "open_items": str(wp5_open_items_path)}, "工作簿构建器和35条WP5治理事项必须存在。")

    output_dir.mkdir(parents=True, exist_ok=True)
    export_dir = run_dir / "10_output/day8_exports"
    export_dir.mkdir(parents=True, exist_ok=True)
    source_map = {
        "standard_31_fields.csv": run_dir / "03_standardized/day4_standard_31_fields.csv",
        "activity_36_fields.csv": run_dir / "05_activity/day5_activity_36_fields.csv",
        "third_party_20_fields.csv": run_dir / "06_third_party_input/day5_third_party_20_fields.csv",
        "quality_issues.csv": run_dir / "04_qc/day5_quality_issue_records.csv",
        "factor_d1_45_fields.csv": run_dir / "07_factor_results/day6_d1_factor_results_45_fields.csv",
        "factor_adaptation_audit.csv": run_dir / "07_factor_results/day6_8_to_45_adaptation_audit.csv",
        "exact_match_d2_57_fields.csv": run_dir / "08_matching/day6_d2_exact_matches_57_fields.csv",
        "route_d3_36_fields.csv": run_dir / "08_matching/day6_d3_exact_routes_36_fields.csv",
        "calculation_d4_48_fields.csv": run_dir / "09_calculation/day7_d4_calculation_48_fields.csv",
        "end_to_end_d5_56_fields.csv": run_dir / "10_output/day7_d5_end_to_end_56_fields.csv",
        "wp5_frozen_lineage_32_fields.csv": run_dir / "10_output/day7_wp5_frozen_lineage_32_fields.csv",
        "demo_extended_lineage_25_fields.csv": run_dir / "10_output/day7_demo_extended_lineage_25_fields.csv",
        "emission_summary_16_fields.csv": run_dir / "09_calculation/day7_emission_summary_16_fields.csv",
        "excluded_records_audit.csv": run_dir / "02_scope_filter/excluded_records_audit.csv",
    }
    exports = {name: _copy_required(source, export_dir / name) for name, source in source_map.items()}

    receipt = _load_json(run_dir / "01_import/file_receipt_report.json")
    receipt_rows: list[dict[str, str]] = []
    _flatten_json("", receipt, receipt_rows)
    receipt_rows.extend([
        {"Property": "day8.run_id", "Value": run_dir.name, "Status": "RECORDED", "Evidence": "CURRENT_RUN"},
        {"Property": "day8.started_at_utc", "Value": started.isoformat(), "Status": "RECORDED", "Evidence": "DAY8_EXPORT"},
    ])
    file_receipt_csv = export_dir / "file_receipt_and_run_metadata.csv"
    _write_csv(file_receipt_csv, receipt_rows, ["Property", "Value", "Status", "Evidence"])

    sheet_inventory = _load_json(run_dir / "01_import/sheet_inventory.json")
    header_detection = {row["sheet_name"]: row for row in _load_json(run_dir / "01_import/header_detection.json")}
    header_rows = []
    for row in sheet_inventory:
        detection = header_detection[row["sheet_name"]]
        header_rows.append({
            "Sheet_Name": row["sheet_name"], "Sheet_State": row["sheet_state"],
            "Physical_Row_Count": row["physical_row_count"], "Column_Count": row["column_count"],
            "Dimension": row["dimension"], "Header_Detected": str(row["header_detected"]).upper(),
            "Header_Row": row["header_row"], "Data_Row_Count": row["data_row_count"],
            "Formula_Count": row["formula_count"], "Merged_Cell_Count": row["merged_cell_count"],
            "Match_Count": detection["match_count"], "Detection_Status": "PASS" if detection["detected"] else "FAIL",
        })
    sheet_header_csv = export_dir / "sheet_and_header_summary.csv"
    _write_csv(sheet_header_csv, header_rows, list(header_rows[0]))

    scope = _load_json(run_dir / "02_scope_filter/day3_scope_summary.json")
    scope_rows: list[dict[str, str]] = []
    _flatten_json("", scope, scope_rows)
    scope_csv = export_dir / "scope_filter_summary.csv"
    _write_csv(scope_csv, scope_rows, ["Property", "Value", "Status", "Evidence"])

    status_rows = _build_status_open_items(
        run_dir / "05_activity/day5_record_open_items.csv",
        run_dir / "05_activity/day5_interface_open_items.csv",
        wp5_open_items_path,
    )
    status_csv = export_dir / "status_and_open_items.csv"
    _write_csv(status_csv, status_rows, STATUS_OPEN_ITEM_FIELDS)
    if sum(row["Layer"] == "WP3_RECORD" for row in status_rows) != 40 or sum(row["Layer"] == "WP3_INTERFACE" for row in status_rows) != 5 or sum(row["Layer"] == "WP5_GOVERNANCE" for row in status_rows) != 35:
        raise _error("OPEN_ITEM_LAYER_COUNTS_INVALID", "40/5/35三类Open Items数量不正确。", status_csv.name, len(status_rows), "三类事项必须分层且不得合并关闭。")

    workbook_path = output_dir / "WP5_Demo_Day8_Result.xlsx"
    verification_path = run_dir / "10_output/day8_workbook_readback.json"
    preview_dir = artifact_work_dir / "previews"
    manifest_path = run_dir / "11_logs/day8_workbook_manifest.json"
    manifest = {
        "workbookTitle": "WP5 Excel自动导入清洗核算Demo — Day 8结果工作簿",
        "workbookSubtitle": f"运行编号：{run_dir.name}。结果来自本次隔离运行；历史因子模拟，禁止生产使用。",
        "expected": {
            "activityTotalKg": day7["activity_total_kg"], "efValue": day7["ef_value"],
            "rawTotal": day7["raw_total_emission_kgco2e"],
            "officialTotal": day7["official_six_decimal_total"],
            "rowDisplayTotal": day7["row_six_decimal_sum"],
            "roundingDifference": day7["rounding_reconciliation_difference"],
        },
        "sheets": [
            {"kind": "overview", "name": "01_运行概览"},
            {"kind": "csv", "name": "02_文件接收", "title": "文件接收与运行元数据", "subtitle": "源文件哈希、隔离副本、运行编号和Day 8开始时间。", "csvPath": str(file_receipt_csv), "tableName": "T02FileReceipt"},
            {"kind": "csv", "name": "03_工作表与表头", "title": "工作表清单与表头识别", "subtitle": "前10行表头检测及两张业务工作表结构。", "csvPath": str(sheet_header_csv), "tableName": "T03SheetHeader"},
            {"kind": "csv", "name": "04_范围筛选", "title": "范围筛选与去向平衡", "subtitle": "合成示例记录均有去向说明。", "csvPath": str(scope_csv), "tableName": "T04Scope"},
            {"kind": "csv", "name": "05_质量问题", "title": "质量问题明细", "subtitle": "上游WARNING证据独立保留，不覆盖WP5阶段状态。", "csvPath": str(exports["quality_issues.csv"]), "tableName": "T05Quality"},
            {"kind": "csv", "name": "06_活动数据", "title": "36字段活动数据", "subtitle": "Decimal活动量，单位kg/year。", "csvPath": str(exports["activity_36_fields.csv"]), "tableName": "T06Activity"},
            {"kind": "csv", "name": "07_第三方输入", "title": "20字段第三方输入", "subtitle": "Record_ID与活动数据一一对应。", "csvPath": str(exports["third_party_20_fields.csv"]), "tableName": "T07ThirdParty"},
            {"kind": "csv", "name": "08_因子结果", "title": "WP5-D1 45字段因子结果", "subtitle": "历史模拟1.250000 kgCO2e/kg，Production_Eligible=FALSE。", "csvPath": str(exports["factor_d1_45_fields.csv"]), "tableName": "T08Factors"},
            {"kind": "csv", "name": "09_匹配与路由", "title": "36字段正式路由", "subtitle": "EXACT_LOCKED，回退和人工复核均为0。", "csvPath": str(exports["route_d3_36_fields.csv"]), "tableName": "T09Route"},
            {"kind": "csv", "name": "10_核算结果", "title": "WP5-D4 48字段自动核算", "subtitle": "Python Decimal逐行核算，保存未舍入值和六位展示值。", "csvPath": str(exports["calculation_d4_48_fields.csv"]), "tableName": "T10Calculation"},
            {"kind": "csv", "name": "11_WP5冻结血缘", "title": "WP5冻结32字段血缘", "subtitle": "严格保持字段名称、顺序和语义，不写入Demo扩展字段。", "csvPath": str(exports["wp5_frozen_lineage_32_fields.csv"]), "tableName": "T11FrozenLineage"},
            {"kind": "csv", "name": "12_汇总与限制", "title": "排放汇总、舍入与使用限制", "subtitle": "先汇总未舍入值，再ROUND_HALF_UP到六位；模拟结果禁止生产使用。", "csvPath": str(exports["emission_summary_16_fields.csv"]), "tableName": "T12Summary"},
            {"kind": "csv", "name": "13_Demo扩展血缘", "title": "Demo扩展25字段血缘", "subtitle": "通过Record_ID连接源文件、工作表、源行、阶段文件和配置快照。", "csvPath": str(exports["demo_extended_lineage_25_fields.csv"]), "tableName": "T13ExtendedLineage"},
            {"kind": "csv", "name": "14_状态与OpenItems", "title": "状态与三类Open Items", "subtitle": "40条记录级、5条接口级、35条WP5治理级事项保持分层且继续OPEN。", "csvPath": str(status_csv), "tableName": "T14OpenItems"},
            {"kind": "csv", "name": "15_D5端到端结果", "title": "56字段端到端结果", "subtitle": "完整结果，后48字段与D4一致。", "csvPath": str(exports["end_to_end_d5_56_fields.csv"]), "tableName": "T15D5Results"},
            {"kind": "csv", "name": "16_D2精确匹配", "title": "57字段精确匹配", "subtitle": "EXACT_RECORD_ID与EXACT_MATCH。", "csvPath": str(exports["exact_match_d2_57_fields.csv"]), "tableName": "T16D2Matches"},
            {"kind": "csv", "name": "17_因子适配审计", "title": "8字段到45字段适配审计", "subtitle": "原始回执、哈希、补全来源和逐条适配状态。", "csvPath": str(exports["factor_adaptation_audit.csv"]), "tableName": "T17FactorAudit"},
            {"kind": "csv", "name": "18_排除审计", "title": "Day 3范围排除审计", "subtitle": "5939条排除记录逐条保留原因，确保6081条记录去向平衡。", "csvPath": str(exports["excluded_records_audit.csv"]), "tableName": "T18Exclusions"},
        ],
    }
    _write_json(manifest_path, manifest)

    artifact_work_dir.mkdir(parents=True, exist_ok=True)
    _ensure_node_modules_junction(artifact_work_dir, node_modules_path)
    builder_copy = artifact_work_dir / "build_day8_workbook.mjs"
    shutil.copy2(builder_script_path, builder_copy)
    completed = subprocess.run(
        [str(node_executable), str(builder_copy), "--manifest", str(manifest_path), "--output", str(workbook_path), "--verification", str(verification_path), "--preview-dir", str(preview_dir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", check=False,
    )
    verification = _load_json(verification_path) if verification_path.is_file() else None
    if completed.returncode != 0 and not verification:
        raise _error("WORKBOOK_BUILD_FAILED", "artifact-tool工作簿构建失败。", builder_copy.name, completed.stderr or completed.stdout, "不得改用openpyxl或其他作者库。")
    if verification is None:
        raise _error("WORKBOOK_VERIFICATION_MISSING", "工作簿构建未生成回读报告。", builder_copy.name, completed.stdout, "artifact-tool必须在导出后重新导入并写出验证报告。")
    if verification.get("status") != "PASS":
        raise _error("WORKBOOK_READBACK_FAILED", "工作簿回读或公式验证失败。", verification_path.name, verification, "18张表、表头、记录数和公式错误扫描必须全部通过。")

    fingerprint_targets = [(f"CSV-{index:02d}", "CSV", path) for index, path in enumerate(sorted(export_dir.glob("*.csv")), start=1)]
    fingerprint_targets.extend([("WORKBOOK-01", "XLSX", workbook_path), ("VERIFY-01", "JSON", verification_path)])
    fingerprint_rows = _fingerprints(fingerprint_targets, run_dir)
    fingerprint_csv = run_dir / "10_output/day8_artifact_fingerprints.csv"
    fingerprint_json = run_dir / "10_output/day8_artifact_fingerprints.json"
    _write_csv(fingerprint_csv, fingerprint_rows, FINGERPRINT_FIELDS)
    _write_json(fingerprint_json, {"algorithm": "SHA-256", "artifact_count": len(fingerprint_rows), "status": "PASS", "artifacts": fingerprint_rows})

    ended = datetime.now(timezone.utc)
    report = {
        "run_id": run_dir.name,
        "status": "PASS",
        "gate_status": "G2_CLI_END_TO_END_PASS",
        "started_at_utc": started.isoformat(),
        "ended_at_utc": ended.isoformat(),
        "elapsed_seconds": round(monotonic() - timer, 3),
        "workbook_path": str(workbook_path),
        "workbook_sha256": verification["workbookSha256"],
        "sheet_count": verification["readbackSheetCount"],
        "formula_error_count": verification["formulaErrorCount"],
        "source_table_readback_equal": verification["allSourceTablesReadBackEqual"],
        "preview_count": verification["previewCount"],
        "csv_export_count": len(list(export_dir.glob("*.csv"))),
        "fingerprinted_artifact_count": len(fingerprint_rows),
        "day7_stage_status": day7["stage_status"],
        "private_data_exported_outside_git": True,
        "github_publication_performed": False,
        "outputs": {
            "workbook": str(workbook_path),
            "csv_directory": str(export_dir),
            "readback_report": str(verification_path),
            "fingerprints_csv": str(fingerprint_csv),
            "fingerprints_json": str(fingerprint_json),
        },
    }
    report_path = run_dir / "10_output/day8_run_report.json"
    _write_json(report_path, report)
    return {**report, "run_directory": str(run_dir), "report_path": str(report_path)}
