from __future__ import annotations

import csv
import hashlib
import json
import math
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from carbon_excel_pipeline.cleaning.raw_cleaner import (
    clean_text,
    contains_marker,
    map_unit_strict,
    parse_strict_positive_decimal,
)
from carbon_excel_pipeline.cleaning.scope_filter import (
    _scope_decision,
    run_day3_scope_and_cleaning,
)
from carbon_excel_pipeline.io.excel_importer import inspect_excel_to_run
from carbon_excel_pipeline.mapping.field_confirmation import confirm_sheet_mapping


ROOT = Path(__file__).resolve().parents[2]
ALIAS_CONFIG = ROOT / "config/import/field_aliases.json"
UNIT_CONFIG = ROOT / "config/cleaning/unit_mappings.json"
TARGET_CATEGORY = "Cell.Polymer Cell.Polymer Cell"
HEADERS = [
    "采购分类",
    "物料描述",
    "求和项:数量",
    "单位",
    "单PCS净重（G/PCS）",
    "采购量（g/年）",
]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


class Day3ScopeCleaningTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.temp_root = Path(self.temp.name)
        self.source = self.temp_root / "synthetic.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Synthetic Input"
        worksheet.append(HEADERS)
        worksheet.append(
            [TARGET_CATEGORY, "P100 synthetic_supplier_a 123456", 10, "PCS", 5.5, 55]
        )
        worksheet.append(
            [TARGET_CATEGORY, "P200 SYNTHETIC_SUPPLIER_A 234567", 20, "PCS", 6, 120]
        )
        worksheet.append([TARGET_CATEGORY, "P300 OTHER 345678", 30, "PCS", 7, 210])
        worksheet.append(
            ["Other.Category", "P400 SYNTHETIC_SUPPLIER_A 456789", 40, "PCS", 8, 320]
        )
        workbook.save(self.source)
        workbook.close()
        self.source_hash = sha256(self.source)
        self.run_result = inspect_excel_to_run(
            self.source,
            run_root=self.temp_root / "runs",
            alias_config_path=ALIAS_CONFIG,
        )
        self.run_dir = Path(self.run_result["run_directory"])
        self.scope_config = self.temp_root / "scope.json"
        self.scope_payload = {
            "config_id": "SYNTHETIC_DAY3_TEST_SCOPE_V1",
            "profile_id": "public_synthetic_profile",
            "target_sheet": "Synthetic Input",
            "target_purchase_category": TARGET_CATEGORY,
            "category_match_method": "CLEANED_EXACT",
            "supplier_markers": ["SYNTHETIC_SUPPLIER_A"],
            "supplier_match_method": "CASE_INSENSITIVE_SUBSTRING",
            "expected_input_records": 4,
            "expected_candidate_records": 2,
            "required_targets": [
                "Purchase_Category",
                "Product_Description",
                "PCS",
                "Source_Unit",
                "Unit_Weight_g_per_PCS",
                "Annual_Purchase_g_per_year",
            ],
            "context_units": {
                "unit_weight_raw": "G/PCS",
                "annual_activity_raw": "g/年",
            },
        }
        self.scope_config.write_text(
            json.dumps(self.scope_payload, ensure_ascii=False), encoding="utf-8"
        )
        self.day3 = run_day3_scope_and_cleaning(
            self.run_dir,
            scope_config_path=self.scope_config,
            unit_config_path=UNIT_CONFIG,
        )

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _read_csv(self, relative: str) -> list[dict[str, str]]:
        with (self.run_dir / relative).open(
            "r", encoding="utf-8-sig", newline=""
        ) as handle:
            return list(csv.DictReader(handle))

    def test_d3_01_confirms_six_unique_field_mappings(self) -> None:
        confirmation = json.loads(
            (self.run_dir / "02_scope_filter/field_mapping_confirmation.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(confirmation["status"], "CONFIRMED")
        self.assertEqual(len(confirmation["confirmed_fields"]), 6)

    def test_d3_02_conflict_requires_manual_confirmation(self) -> None:
        report = [
            {
                "sheet_name": "Synthetic Input",
                "header_row": 1,
                "mapping_preview": [
                    {
                        "target_field": "PCS",
                        "status": "CONFLICT",
                        "matched_columns": [],
                    }
                ],
            }
        ]
        result = confirm_sheet_mapping(
            report, target_sheet="Synthetic Input", required_targets=["PCS"]
        )
        self.assertEqual(result["status"], "MANUAL_CONFIRMATION_REQUIRED")
        self.assertTrue(result["requires_manual_confirmation"])

    def test_d3_03_accounts_for_every_target_sheet_record(self) -> None:
        self.assertEqual(self.day3["input_records"], 4)
        self.assertTrue(self.day3["destination_balance"])

    def test_d3_04_produces_expected_candidate_count(self) -> None:
        self.assertEqual(self.day3["candidate_records"], 2)
        self.assertEqual(len(self._read_csv("02_scope_filter/candidate_records.csv")), 2)

    def test_d3_05_produces_complete_exclusion_audit(self) -> None:
        rows = self._read_csv("02_scope_filter/excluded_records_audit.csv")
        self.assertEqual(len(rows), 2)
        self.assertTrue(all(row["Source_Row"] for row in rows))

    def test_d3_06_applies_exclusion_reason_precedence(self) -> None:
        self.assertEqual(
            self.day3["exclusion_reason_counts"],
            {"CATEGORY_OUT_OF_SCOPE": 1, "SUPPLIER_MARKER_NOT_FOUND": 1},
        )

    def test_d3_07_supplier_text_change_changes_scope_result(self) -> None:
        included = _scope_decision(
            category_clean=TARGET_CATEGORY,
            description_clean="P100 synthetic_supplier_a 123456",
            config=self.scope_payload,
        )[0]
        excluded = _scope_decision(
            category_clean=TARGET_CATEGORY,
            description_clean="P100 OTHER 123456",
            config=self.scope_payload,
        )[0]
        self.assertTrue(included)
        self.assertFalse(excluded)

    def test_d3_08_normalizes_text_without_overwriting_raw(self) -> None:
        self.assertEqual(clean_text("  SYNA　 物料  "), "SYNA 物料")
        rows = self._read_csv("03_standardized/day3_cleaned_candidates.csv")
        self.assertIn("Product_Description_Raw", rows[0])
        self.assertIn("Product_Description_Clean", rows[0])

    def test_d3_09_parses_positive_decimal_canonically(self) -> None:
        result = parse_strict_positive_decimal(
            Decimal("55.2000"), field_code="UNIT_WEIGHT"
        )
        self.assertTrue(result.is_valid)
        self.assertEqual(result.cleaned_value, "55.2")

    def test_d3_10_rejects_numeric_text(self) -> None:
        result = parse_strict_positive_decimal("10", field_code="PCS")
        self.assertEqual(result.issue_code, "PCS_TEXT_NOT_ALLOWED")

    def test_d3_11_marks_zero_negative_nan_and_infinity(self) -> None:
        cases = [
            (0, "PCS_NOT_STRICTLY_POSITIVE"),
            (-1, "PCS_NOT_STRICTLY_POSITIVE"),
            (math.nan, "PCS_NON_FINITE"),
            (math.inf, "PCS_NON_FINITE"),
        ]
        for value, expected in cases:
            with self.subTest(value=value):
                self.assertEqual(
                    parse_strict_positive_decimal(value, field_code="PCS").issue_code,
                    expected,
                )

    def test_d3_12_requires_integer_pcs(self) -> None:
        result = parse_strict_positive_decimal(
            1.5, field_code="PCS", integer_required=True
        )
        self.assertEqual(result.issue_code, "PCS_NOT_INTEGER")

    def test_d3_13_maps_only_explicit_machine_units(self) -> None:
        pcs = map_unit_strict("PCS", field_code="PCS", mappings={"PCS": "PCS"})
        weight = map_unit_strict(
            "G/PCS", field_code="UNIT_WEIGHT", mappings={"G/PCS": "g/PCS"}
        )
        annual = map_unit_strict(
            "g/年", field_code="ANNUAL_ACTIVITY", mappings={"g/年": "g/year"}
        )
        self.assertEqual(
            [pcs.cleaned_unit, weight.cleaned_unit, annual.cleaned_unit],
            ["PCS", "g/PCS", "g/year"],
        )

    def test_d3_14_does_not_trim_or_casefold_machine_units(self) -> None:
        mappings = {"PCS": "PCS"}
        self.assertFalse(
            map_unit_strict("PCS ", field_code="PCS", mappings=mappings).is_valid
        )
        self.assertFalse(
            map_unit_strict("pcs", field_code="PCS", mappings=mappings).is_valid
        )

    def test_d3_15_preserves_raw_and_clean_numeric_values(self) -> None:
        rows = self._read_csv("03_standardized/day3_cleaned_candidates.csv")
        self.assertEqual(rows[0]["PCS_Raw"], "10")
        self.assertEqual(rows[0]["PCS_Clean"], "10")
        self.assertEqual(rows[0]["Cleaning_Status"], "PASS")

    def test_d3_16_writes_all_outputs_without_changing_source(self) -> None:
        expected = [
            "02_scope_filter/field_mapping_confirmation.json",
            "02_scope_filter/candidate_records.csv",
            "02_scope_filter/excluded_records_audit.csv",
            "02_scope_filter/day3_scope_summary.json",
            "03_standardized/day3_cleaned_candidates.csv",
            "03_standardized/cleaning_issue_records.csv",
            "03_standardized/day3_cleaning_summary.json",
        ]
        self.assertTrue(all((self.run_dir / item).is_file() for item in expected))
        self.assertEqual(sha256(self.source), self.source_hash)
        self.assertTrue(contains_marker("abc synthetic_supplier_a xyz", "SYNTHETIC_SUPPLIER_A"))


if __name__ == "__main__":
    unittest.main()
