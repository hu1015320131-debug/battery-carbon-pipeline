from __future__ import annotations

import unittest
from decimal import Decimal

from carbon_excel_pipeline.wp6_6.pipeline import (
    WP66AnalysisError,
    build_scenario_analysis,
)


def _row(year: str, record_id: str, activity: str, ef: str) -> dict[str, str]:
    emission = Decimal(activity) * Decimal(ef)
    return {
        "Year": year,
        "Record_ID": record_id,
        "Independent_Activity_kg": activity,
        "Independent_EF": ef,
        "Independent_Emission_kgCO2e": format(emission, "f"),
        "Overall_Validation_Status": "INDEPENDENT_CALCULATION_PASS",
        "Boundary_Ready": "TRUE",
        "Lineage_Validation_Status": "PASS",
        "Activity_Validation_Status": "PASS_EXACT",
        "EF_Validation_Status": "PASS_EXACT",
        "Emission_Validation_Status": "PASS_EXACT",
    }


class WP66ScenarioTests(unittest.TestCase):
    def setUp(self) -> None:
        self.rows_2024 = [
            _row("2024", "2024-DY2-SYNA-DX000001", "10", "2"),
            _row("2024", "2024-DY2-SYNA-DX000002", "20", "2"),
        ]
        self.rows_2025 = [
            _row("2025", "2025-DY2-SYNA-DX000001", "40", "1.5"),
            _row("2025", "2025-DY2-SYNA-DX000002", "60", "1.5"),
        ]
        self.analysis = build_scenario_analysis(self.rows_2024, self.rows_2025)

    def test_scenario_a(self) -> None:
        self.assertEqual(
            Decimal(self.analysis["scenario_abcd"]["A"]["Emission_kgCO2e"]),
            Decimal("60"),
        )

    def test_scenario_b(self) -> None:
        self.assertEqual(
            Decimal(self.analysis["scenario_abcd"]["B"]["Emission_kgCO2e"]),
            Decimal("45.0"),
        )

    def test_scenario_c(self) -> None:
        self.assertEqual(
            Decimal(self.analysis["scenario_abcd"]["C"]["Emission_kgCO2e"]),
            Decimal("200"),
        )

    def test_scenario_d(self) -> None:
        self.assertEqual(
            Decimal(self.analysis["scenario_abcd"]["D"]["Emission_kgCO2e"]),
            Decimal("150.0"),
        )

    def test_factor_effects(self) -> None:
        factor = self.analysis["factor_effect"]
        self.assertEqual(
            Decimal(factor["Factor_Effect_At_2024_Activity"]["Effect_kgCO2e"]),
            Decimal("-15"),
        )
        self.assertEqual(
            Decimal(factor["Factor_Effect_At_2025_Activity"]["Effect_kgCO2e"]),
            Decimal("-50"),
        )

    def test_activity_scope_effects(self) -> None:
        effect = self.analysis["activity_scope_effect"]
        self.assertEqual(
            Decimal(
                effect["Activity_Scope_Combined_Effect_At_2024_EF"][
                    "Effect_kgCO2e"
                ]
            ),
            Decimal("140"),
        )
        self.assertEqual(
            Decimal(
                effect["Activity_Scope_Combined_Effect_At_2025_EF"][
                    "Effect_kgCO2e"
                ]
            ),
            Decimal("105"),
        )

    def test_symmetric_decomposition_reconciles_exactly(self) -> None:
        symmetric = self.analysis["symmetric_decomposition"]
        factor = Decimal(
            symmetric["Symmetric_Factor_Contribution"]["Effect_kgCO2e"]
        )
        activity_scope = Decimal(
            symmetric["Symmetric_Activity_Scope_Contribution"]["Effect_kgCO2e"]
        )
        observed = Decimal(
            symmetric["Observed_Cross_Year_Difference"]["Effect_kgCO2e"]
        )
        self.assertEqual(factor + activity_scope, observed)
        self.assertEqual(symmetric["Reconciliation_Status"], "PASS_EXACT")

    def test_scope_safety_never_matches_same_suffix_across_years(self) -> None:
        self.assertFalse(
            self.analysis["scope_safety"]["cross_year_record_matching_performed"]
        )
        self.assertFalse(
            self.analysis["scope_safety"]["record_id_suffix_matching_performed"]
        )
        rows_2024 = self.analysis["record_counterfactuals"]["2024"]
        rows_2025 = self.analysis["record_counterfactuals"]["2025"]
        self.assertTrue(all(row["Year"] == "2024" for row in rows_2024))
        self.assertTrue(all(row["Year"] == "2025" for row in rows_2025))

    def test_nonvalidated_record_is_explicitly_excluded(self) -> None:
        blocked = dict(self.rows_2024[0])
        blocked["Overall_Validation_Status"] = "INDEPENDENT_VALIDATION_FAIL"
        analysis = build_scenario_analysis(
            [blocked, self.rows_2024[1]], self.rows_2025
        )
        self.assertEqual(analysis["status"], "PASS_WITH_WARNING")
        self.assertEqual(analysis["coverage"]["Excluded_Record_Count"], 1)
        self.assertIn(
            "WP6_5_INDEPENDENT_VALIDATION_NOT_PASS",
            analysis["excluded_records"][0]["Reason_Codes"],
        )

    def test_multiple_ef_values_block_single_factor_model(self) -> None:
        different = dict(self.rows_2024[1])
        different["Independent_EF"] = "3"
        different["Independent_Emission_kgCO2e"] = "60"
        with self.assertRaisesRegex(
            WP66AnalysisError, "MULTIPLE_EF_REQUIRES_MODEL_REASSESSMENT"
        ):
            build_scenario_analysis([self.rows_2024[0], different], self.rows_2025)

    def test_record_level_factor_percent_is_common_but_impact_is_not(self) -> None:
        rows = self.analysis["record_counterfactuals"]["2024"]
        self.assertEqual(rows[0]["Factor_Impact_Percent"], rows[1]["Factor_Impact_Percent"])
        self.assertNotEqual(rows[0]["Factor_Impact_kgCO2e"], rows[1]["Factor_Impact_kgCO2e"])


if __name__ == "__main__":
    unittest.main()
