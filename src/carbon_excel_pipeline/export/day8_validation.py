"""Independent, read-only validation for the Day 8 XLSX delivery."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree


EXPECTED_SHEET_NAMES = [
    "01_运行概览",
    "02_文件接收",
    "03_工作表与表头",
    "04_范围筛选",
    "05_质量问题",
    "06_活动数据",
    "07_第三方输入",
    "08_因子结果",
    "09_匹配与路由",
    "10_核算结果",
    "11_WP5冻结血缘",
    "12_汇总与限制",
    "13_Demo扩展血缘",
    "14_状态与OpenItems",
    "15_D5端到端结果",
    "16_D2精确匹配",
    "17_因子适配审计",
    "18_排除审计",
]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def workbook_sheet_names(path: Path) -> list[str]:
    """Read the ordered sheet list directly from the XLSX package."""
    with zipfile.ZipFile(path) as archive:
        workbook_xml = archive.read("xl/workbook.xml")
    root = ElementTree.fromstring(workbook_xml)
    namespace = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return [node.attrib["name"] for node in root.findall("m:sheets/m:sheet", namespace)]


def validate_day8_delivery(report_path: Path) -> dict[str, Any]:
    """Validate the G2 report, XLSX package, readback report and fingerprints."""
    report_path = report_path.expanduser().resolve()
    errors: list[str] = []
    if not report_path.is_file():
        return {"status": "FAIL", "errors": [f"Missing report: {report_path}"], "checks": {}}

    report = json.loads(report_path.read_text(encoding="utf-8"))
    checks: dict[str, bool] = {}

    checks["g2_gate"] = report.get("status") == "PASS" and report.get("gate_status") == "G2_CLI_END_TO_END_PASS"
    checks["github_not_used"] = report.get("github_publication_performed") is False
    checks["private_output_outside_git"] = report.get("private_data_exported_outside_git") is True

    workbook_path = Path(str(report.get("workbook_path", "")))
    checks["workbook_exists"] = workbook_path.is_file()
    if checks["workbook_exists"]:
        checks["workbook_sha256"] = sha256_file(workbook_path) == report.get("workbook_sha256")
        try:
            checks["sheet_contract"] = workbook_sheet_names(workbook_path) == EXPECTED_SHEET_NAMES
        except (KeyError, OSError, zipfile.BadZipFile, ElementTree.ParseError):
            checks["sheet_contract"] = False
    else:
        checks["workbook_sha256"] = False
        checks["sheet_contract"] = False

    outputs = report.get("outputs", {})
    readback_path = Path(str(outputs.get("readback_report", "")))
    checks["readback_exists"] = readback_path.is_file()
    if checks["readback_exists"]:
        readback = json.loads(readback_path.read_text(encoding="utf-8"))
        checks["readback_pass"] = readback.get("status") == "PASS"
        checks["readback_sheet_count"] = readback.get("readbackSheetCount") == len(EXPECTED_SHEET_NAMES)
        checks["formula_errors_zero"] = readback.get("formulaErrorCount") == 0
        checks["source_tables_equal"] = readback.get("allSourceTablesReadBackEqual") is True
        checks["all_previews_created"] = readback.get("previewCount") == len(EXPECTED_SHEET_NAMES)
    else:
        for name in (
            "readback_pass",
            "readback_sheet_count",
            "formula_errors_zero",
            "source_tables_equal",
            "all_previews_created",
        ):
            checks[name] = False

    csv_directory = Path(str(outputs.get("csv_directory", "")))
    checks["csv_export_count"] = csv_directory.is_dir() and len(list(csv_directory.glob("*.csv"))) == 18

    fingerprints_path = Path(str(outputs.get("fingerprints_json", "")))
    checks["fingerprints_exist"] = fingerprints_path.is_file()
    if checks["fingerprints_exist"]:
        fingerprints = json.loads(fingerprints_path.read_text(encoding="utf-8"))
        artifacts = fingerprints.get("artifacts", [])
        checks["fingerprint_count"] = fingerprints.get("artifact_count") == 20 and len(artifacts) == 20
        workbook_rows = [row for row in artifacts if row.get("Artifact_ID") == "WORKBOOK-01"]
        checks["workbook_fingerprint"] = (
            len(workbook_rows) == 1
            and workbook_rows[0].get("SHA256") == report.get("workbook_sha256")
        )
    else:
        checks["fingerprint_count"] = False
        checks["workbook_fingerprint"] = False

    errors.extend(name for name, passed in checks.items() if not passed)
    return {
        "status": "PASS" if not errors else "FAIL",
        "gate_status": "G2_CLI_END_TO_END_PASS" if not errors else "G2_BLOCKED",
        "report_path": str(report_path),
        "workbook_path": str(workbook_path),
        "checks": checks,
        "errors": errors,
    }
