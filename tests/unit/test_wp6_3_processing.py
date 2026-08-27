from __future__ import annotations

import unittest
from decimal import Decimal

from carbon_excel_pipeline.wp6_3.processing import (
    adapt_direct_mass,
    adapt_historical_ef,
    assign_2024_record_ids,
    calculate_and_validate,
    controlled_forward_fill,
    filter_2024_boundary,
)


class WP63ProcessingTests(unittest.TestCase):
    def test_controlled_forward_fill_preserves_raw_blanks_and_source_rows(self) -> None:
        records = [
            {
                "Source_Row": 8,
                "values": {
                    "Business_Unit": "二部",
                    "Purchase_Type": "合成",
                    "Purchase_Category": "电芯.聚合物",
                    "Product_Description": "first",
                    "Reported_Activity_Value": 1,
                    "EF_Value": 2,
                },
            },
            {
                "Source_Row": 9,
                "values": {
                    "Business_Unit": None,
                    "Purchase_Type": None,
                    "Purchase_Category": None,
                    "Product_Description": None,
                    "Reported_Activity_Value": None,
                    "EF_Value": None,
                },
            },
        ]
        result = controlled_forward_fill(records)
        self.assertEqual(result[1]["Source_Row"], 9)
        self.assertEqual(result[1]["values"]["Business_Unit"], "二部")
        self.assertIsNone(result[1]["original_context_values"]["Business_Unit"])
        self.assertEqual(
            result[1]["context_value_sources"]["Business_Unit"],
            "CONTROLLED_FORWARD_FILL",
        )
        self.assertIsNone(result[1]["values"]["Product_Description"])
        self.assertIsNone(result[1]["values"]["Reported_Activity_Value"])
        self.assertIsNone(result[1]["values"]["EF_Value"])

    def test_boundary_uses_fields_not_fixed_row_numbers(self) -> None:
        records = [
            {
                "Source_Row": 900,
                "values": {
                    "Business_Unit": "二部",
                    "Purchase_Category": "电芯.聚合物电芯",
                    "Product_Description": "demo SYNA cell",
                },
            },
            {
                "Source_Row": 4,
                "values": {
                    "Business_Unit": "二部",
                    "Purchase_Category": "电芯.聚合物电芯",
                    "Product_Description": "other cell",
                },
            },
        ]
        selected, audit = filter_2024_boundary(
            records, business_unit="二部", category_root="电芯", product_marker="SYNA"
        )
        self.assertEqual([item["Source_Row"] for item in selected], [900])
        self.assertEqual([item["Selected_Count"] for item in audit], [2, 2, 1])

    def test_record_ids_are_stable_by_source_row(self) -> None:
        records = [{"Source_Row": 20}, {"Source_Row": 10}]
        first = assign_2024_record_ids(records)
        second = assign_2024_record_ids(list(reversed(records)))
        expected = [(10, "2024-DY2-SYNA-DX000001"), (20, "2024-DY2-SYNA-DX000002")]
        self.assertEqual(
            [(item["Source_Row"], item["Record_ID"]) for item in first], expected
        )
        self.assertEqual(
            [(item["Source_Row"], item["Record_ID"]) for item in second], expected
        )

    def test_direct_mass_conversions(self) -> None:
        self.assertEqual(adapt_direct_mass(1, "t/year")["activity_kg"], Decimal("1000"))
        self.assertEqual(adapt_direct_mass(1, "kg/year")["activity_kg"], Decimal("1"))
        self.assertEqual(adapt_direct_mass(1, "g/year")["activity_kg"], Decimal("0.001"))

    def test_activity_negative_cases_are_record_level(self) -> None:
        self.assertIn("ACTIVITY_MISSING", adapt_direct_mass(None, "kg/year")["blocking"])
        self.assertIn("ACTIVITY_NON_NUMERIC", adapt_direct_mass("x", "kg/year")["blocking"])
        self.assertIn("ZERO_ACTIVITY", adapt_direct_mass(0, "kg/year")["warnings"])
        self.assertIn("ACTIVITY_NEGATIVE", adapt_direct_mass(-1, "kg/year")["blocking"])
        self.assertIn("ACTIVITY_UNIT_UNSUPPORTED", adapt_direct_mass(1, "lb/year")["blocking"])

    def test_historical_ef_is_preserved_and_not_replaced(self) -> None:
        result = adapt_historical_ef("14.25", "kgCO2e/kg", "历史来源")
        self.assertEqual(result["normalized_value"], Decimal("14.25"))
        self.assertEqual(result["source"], "历史来源")

    def test_decimal_calculation_reports_actual_historical_difference(self) -> None:
        result = calculate_and_validate(
            activity_kg=Decimal("1000"),
            ef_value=Decimal("2.5"),
            historical_tco2e=Decimal("2.4999999995"),
            difference_threshold_tco2e=Decimal("0.000000001"),
        )
        self.assertEqual(result["emission_t"], Decimal("2.5"))
        self.assertEqual(result["difference_t"], Decimal("0.0000000005"))
        self.assertEqual(
            result["validation_status"],
            "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE",
        )


if __name__ == "__main__":
    unittest.main()
