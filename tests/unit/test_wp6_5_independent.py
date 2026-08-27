from __future__ import annotations

import unittest
from decimal import Decimal
from pathlib import Path

from carbon_excel_pipeline.wp6_5.independent import (
    calculate_direct_mass_activity,
    calculate_emission,
    calculate_pcs_weight_activity,
    compare_calculation,
)
from carbon_excel_pipeline.wp6_5.pipeline import verify_validator_independence


class WP65IndependentTests(unittest.TestCase):
    def test_validator_has_no_production_calculation_import(self) -> None:
        source = (
            Path(__file__).resolve().parents[2]
            / "src/carbon_excel_pipeline/wp6_5/independent.py"
        )
        audit = verify_validator_independence(source)
        self.assertEqual(audit["status"], "PASS")
        self.assertFalse(audit["production_calculation_imported"])

    def test_2024_supported_mass_units(self) -> None:
        self.assertEqual(calculate_direct_mass_activity("1", "t/year"), Decimal("1000"))
        self.assertEqual(calculate_direct_mass_activity("1", "kg/year"), Decimal("1"))
        self.assertEqual(calculate_direct_mass_activity("1000", "g/year"), Decimal("1.000"))

    def test_2025_pcs_times_weight(self) -> None:
        self.assertEqual(
            calculate_pcs_weight_activity("122380", "55.2", "g/PCS"),
            Decimal("6755.376"),
        )

    def test_decimal_emission_is_independent(self) -> None:
        self.assertEqual(
            calculate_emission("6755.376", "1.250000", "kgCO2e/kg"),
            Decimal("8444.220000000"),
        )

    @staticmethod
    def _comparison(**overrides: object) -> dict[str, str]:
        values: dict[str, object] = {
            "main_activity_kg": "10",
            "independent_activity_kg": "10",
            "main_ef": "2",
            "independent_ef": "2",
            "main_ef_unit": "kgCO2e/kg",
            "independent_ef_unit": "kgCO2e/kg",
            "main_emission_kgco2e": "20",
            "independent_emission_kgco2e": "20",
        }
        values.update(overrides)
        return compare_calculation(**values)

    def test_intentional_activity_error(self) -> None:
        result = self._comparison(main_activity_kg="11")
        self.assertIn("ACTIVITY_CONVERSION_MISMATCH", result["Reason_Codes"])
        self.assertEqual(result["Overall_Validation_Status"], "INDEPENDENT_VALIDATION_FAIL")

    def test_intentional_ef_error(self) -> None:
        result = self._comparison(main_ef="3")
        self.assertIn("EF_VALUE_MISMATCH", result["Reason_Codes"])

    def test_intentional_ef_unit_error(self) -> None:
        result = self._comparison(main_ef_unit="kgCO2e/t")
        self.assertIn("EF_UNIT_MISMATCH", result["Reason_Codes"])

    def test_activity_input_mismatch(self) -> None:
        result = self._comparison(activity_input_match=False)
        self.assertIn("ACTIVITY_INPUT_MISMATCH", result["Reason_Codes"])

    def test_intentional_emission_error(self) -> None:
        result = self._comparison(main_emission_kgco2e="21")
        self.assertIn("EMISSION_CALCULATION_MISMATCH", result["Reason_Codes"])

    def test_lineage_mismatch_even_when_values_match(self) -> None:
        result = self._comparison(lineage_match=False)
        self.assertIn("LINEAGE_MISMATCH", result["Reason_Codes"])
        self.assertEqual(result["Lineage_Validation_Status"], "FAIL")

    def test_governance_warning_is_not_a_calculation_input(self) -> None:
        result = self._comparison()
        self.assertEqual(result["Overall_Validation_Status"], "INDEPENDENT_CALCULATION_PASS")

    def test_nonzero_raw_difference_is_not_hidden_by_display_match(self) -> None:
        result = self._comparison(main_emission_kgco2e="20.0000001")
        self.assertIn("ROUNDING_ONLY_DIFFERENCE", result["Reason_Codes"])
        self.assertEqual(result["Overall_Validation_Status"], "INDEPENDENT_VALIDATION_FAIL")


if __name__ == "__main__":
    unittest.main()
