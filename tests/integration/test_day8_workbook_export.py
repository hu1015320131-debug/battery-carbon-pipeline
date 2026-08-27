from __future__ import annotations

import json
import tempfile
import unittest
import zipfile
from copy import deepcopy
from pathlib import Path

from carbon_excel_pipeline.export.day8_validation import (
    EXPECTED_SHEET_NAMES,
    sha256_file,
    validate_day8_delivery,
    workbook_sheet_names,
)


class Day8WorkbookExportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.report_path = self._make_delivery()

    def tearDown(self) -> None:
        self.temp.cleanup()

    @staticmethod
    def _workbook_xml(names: list[str]) -> str:
        sheets = "".join(
            f'<sheet name="{name}" sheetId="{index}" r:id="rId{index}"/>'
            for index, name in enumerate(names, start=1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheets}</sheets></workbook>"
        )

    def _write_workbook(self, names: list[str]) -> Path:
        path = self.root / "WP5_Demo_Day8_Result.xlsx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("xl/workbook.xml", self._workbook_xml(names))
        return path

    def _make_delivery(self) -> Path:
        workbook = self._write_workbook(EXPECTED_SHEET_NAMES)
        csv_dir = self.root / "day8_exports"
        csv_dir.mkdir()
        for index in range(1, 19):
            (csv_dir / f"export_{index:02d}.csv").write_text("Field\nValue\n", encoding="utf-8")

        readback = {
            "status": "PASS",
            "readbackSheetCount": 18,
            "formulaErrorCount": 0,
            "allSourceTablesReadBackEqual": True,
            "previewCount": 18,
        }
        readback_path = self.root / "day8_workbook_readback.json"
        readback_path.write_text(json.dumps(readback), encoding="utf-8")

        workbook_hash = sha256_file(workbook)
        artifacts = [
            {
                "Artifact_ID": f"CSV-{index:02d}",
                "SHA256": "A" * 64,
            }
            for index in range(1, 19)
        ]
        artifacts.extend(
            [
                {"Artifact_ID": "WORKBOOK-01", "SHA256": workbook_hash},
                {"Artifact_ID": "VERIFY-01", "SHA256": sha256_file(readback_path)},
            ]
        )
        fingerprints_path = self.root / "day8_artifact_fingerprints.json"
        fingerprints_path.write_text(
            json.dumps({"status": "PASS", "artifact_count": 20, "artifacts": artifacts}),
            encoding="utf-8",
        )
        report = {
            "status": "PASS",
            "gate_status": "G2_CLI_END_TO_END_PASS",
            "workbook_path": str(workbook),
            "workbook_sha256": workbook_hash,
            "private_data_exported_outside_git": True,
            "github_publication_performed": False,
            "outputs": {
                "readback_report": str(readback_path),
                "csv_directory": str(csv_dir),
                "fingerprints_json": str(fingerprints_path),
            },
        }
        report_path = self.root / "day8_run_report.json"
        report_path.write_text(json.dumps(report), encoding="utf-8")
        return report_path

    def _mutate_report(self, change) -> None:
        report = json.loads(self.report_path.read_text(encoding="utf-8"))
        change(report)
        self.report_path.write_text(json.dumps(report), encoding="utf-8")

    def _mutate_readback(self, change) -> None:
        report = json.loads(self.report_path.read_text(encoding="utf-8"))
        path = Path(report["outputs"]["readback_report"])
        payload = json.loads(path.read_text(encoding="utf-8"))
        change(payload)
        path.write_text(json.dumps(payload), encoding="utf-8")

    def test_d8_01_sheet_contract_has_18_unique_ordered_names(self) -> None:
        self.assertEqual(len(EXPECTED_SHEET_NAMES), 18)
        self.assertEqual(len(set(EXPECTED_SHEET_NAMES)), 18)
        self.assertEqual((EXPECTED_SHEET_NAMES[0], EXPECTED_SHEET_NAMES[-1]), ("01_运行概览", "18_排除审计"))

    def test_d8_02_xlsx_sheet_names_are_read_from_package(self) -> None:
        self.assertEqual(workbook_sheet_names(self.root / "WP5_Demo_Day8_Result.xlsx"), EXPECTED_SHEET_NAMES)

    def test_d8_03_sha256_changes_when_workbook_changes(self) -> None:
        before = sha256_file(self.root / "WP5_Demo_Day8_Result.xlsx")
        changed = self.root / "changed.xlsx"
        changed.write_bytes((self.root / "WP5_Demo_Day8_Result.xlsx").read_bytes() + b"changed")
        self.assertNotEqual(before, sha256_file(changed))

    def test_d8_04_complete_delivery_passes_g2(self) -> None:
        result = validate_day8_delivery(self.report_path)
        self.assertEqual((result["status"], result["gate_status"]), ("PASS", "G2_CLI_END_TO_END_PASS"), result)

    def test_d8_05_non_g2_report_is_blocked(self) -> None:
        self._mutate_report(lambda report: report.update(gate_status="G1A_UPSTREAM_REBUILD_RECONCILED"))
        self.assertIn("g2_gate", validate_day8_delivery(self.report_path)["errors"])

    def test_d8_06_wrong_sheet_order_is_blocked(self) -> None:
        names = deepcopy(EXPECTED_SHEET_NAMES)
        names[0], names[1] = names[1], names[0]
        workbook = self._write_workbook(names)
        self._mutate_report(lambda report: report.update(workbook_sha256=sha256_file(workbook)))
        self.assertIn("sheet_contract", validate_day8_delivery(self.report_path)["errors"])

    def test_d8_07_workbook_hash_mismatch_is_blocked(self) -> None:
        self._mutate_report(lambda report: report.update(workbook_sha256="0" * 64))
        self.assertIn("workbook_sha256", validate_day8_delivery(self.report_path)["errors"])

    def test_d8_08_formula_errors_are_blocked(self) -> None:
        self._mutate_readback(lambda readback: readback.update(formulaErrorCount=1))
        self.assertIn("formula_errors_zero", validate_day8_delivery(self.report_path)["errors"])

    def test_d8_09_source_table_readback_difference_is_blocked(self) -> None:
        self._mutate_readback(lambda readback: readback.update(allSourceTablesReadBackEqual=False))
        self.assertIn("source_tables_equal", validate_day8_delivery(self.report_path)["errors"])

    def test_d8_10_missing_csv_export_is_blocked(self) -> None:
        (self.root / "day8_exports/export_18.csv").unlink()
        self.assertIn("csv_export_count", validate_day8_delivery(self.report_path)["errors"])

    def test_d8_11_wrong_fingerprint_count_is_blocked(self) -> None:
        path = self.root / "day8_artifact_fingerprints.json"
        payload = json.loads(path.read_text(encoding="utf-8"))
        payload["artifact_count"] = 19
        path.write_text(json.dumps(payload), encoding="utf-8")
        self.assertIn("fingerprint_count", validate_day8_delivery(self.report_path)["errors"])

    def test_d8_12_github_publication_flag_is_blocked(self) -> None:
        self._mutate_report(lambda report: report.update(github_publication_performed=True))
        self.assertIn("github_not_used", validate_day8_delivery(self.report_path)["errors"])


if __name__ == "__main__":
    unittest.main()
