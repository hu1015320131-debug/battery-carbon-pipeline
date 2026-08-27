from __future__ import annotations

import unittest
from decimal import Decimal

from carbon_excel_pipeline.wp6_7 import (
    build_data_quality_scorecard,
    build_dimension_availability,
    build_issue_register,
    build_lineage_quality_summary,
    build_management_summary,
    build_top_emission_contributors,
    build_top_factor_impact,
)


def _row(
    record_id: str,
    *,
    governance: str = "PASS",
    chemistry: str = "LFP",
    emission: str = "20",
    activity: str = "10",
    issue: str = "",
) -> dict[str, str]:
    year = record_id[:4]
    return {
        "Year": year,
        "Record_ID": record_id,
        "Source_File": "source.xlsx",
        "Source_SHA256": "A" * 64,
        "Source_Sheet": "Sheet1",
        "Source_Row": "2",
        "Business_Unit": "BU",
        "Purchase_Category": "Cell",
        "Product_Description": "Product",
        "Chemistry": chemistry,
        "Supplier": "Supplier",
        "Project": "Project",
        "Model": "Model",
        "Activity_kg": activity,
        "Activity_Unit": "kg/year",
        "Activity_Method": "DIRECT_REPORTED_MASS",
        "EF_Value": "2",
        "EF_Unit": "kgCO2e/kg",
        "EF_Source": "SYNTHETIC_FIXTURE",
        "EF_Usage": "HISTORICAL_REPRODUCTION",
        "Emission_kgCO2e": emission,
        "Calculation_QC": "PASS",
        "Governance_QC": governance,
        "Boundary_Ready": "TRUE",
        "Overall_Validation_Status": "INDEPENDENT_CALCULATION_PASS",
        "Calculation_Issue_Codes": "",
        "Governance_Issue_Codes": issue,
        "Boundary_Issue_Codes": "",
        "Run_ID": "RUN-1",
    }


class WP67AnalysisTests(unittest.TestCase):
    def test_quality_denominators_are_explicit(self) -> None:
        scorecard = build_data_quality_scorecard(
            [_row("2024-A"), _row("2024-B")], "2024"
        )
        for metric in scorecard["metrics"]:
            self.assertEqual(metric["Numerator"], 2)
            self.assertEqual(metric["Denominator"], 2)
            self.assertEqual(metric["Applicable_Record_Count"], 2)
            self.assertEqual(metric["Not_Applicable_Count"], 0)
        self.assertFalse(scorecard["Composite_Quality_Score_Created"])

    def test_governance_warning_does_not_reduce_calculation_readiness(self) -> None:
        row = _row(
            "2024-A",
            governance="WARNING",
            chemistry="",
            issue="CHEMISTRY_MISSING",
        )
        metrics = {
            item["Metric"]: item
            for item in build_data_quality_scorecard([row], "2024")["metrics"]
        }
        self.assertEqual(metrics["Calculation_Readiness"]["Numerator"], 1)
        self.assertEqual(metrics["Governance_Field_Completeness"]["Numerator"], 0)

    def test_missing_dimension_is_not_ready_or_grouped_as_zero(self) -> None:
        rows = [_row("2025-A"), _row("2025-B", chemistry="UNKNOWN")]
        availability = build_dimension_availability(rows, "2025")
        chemistry = next(row for row in availability if row["Dimension"] == "Chemistry")
        self.assertEqual(chemistry["Available_Record_Count"], 1)
        self.assertEqual(chemistry["Missing_Record_Count"], 1)
        self.assertEqual(chemistry["Analysis_Ready"], "FALSE")
        summary = build_management_summary(rows, "2025", availability)
        self.assertFalse(any(row["Dimension"] == "Chemistry" for row in summary))
        self.assertFalse(any(row["Dimension_Value"] == "UNKNOWN" for row in summary))

    def test_absent_dimension_is_data_not_available(self) -> None:
        row = _row("2024-A")
        del row["Chemistry"]
        chemistry = next(
            item
            for item in build_dimension_availability([row], "2024")
            if item["Dimension"] == "Chemistry"
        )
        self.assertEqual(chemistry["Field_Available"], "FALSE")
        self.assertEqual(chemistry["Availability_Status"], "DATA_NOT_AVAILABLE")

    def test_issue_priority_is_rule_based(self) -> None:
        calculation = _row("2025-A")
        calculation["Activity_kg"] = ""
        governance = _row(
            "2025-B",
            governance="WARNING",
            chemistry="",
            issue="CHEMISTRY_MISSING",
        )
        register, _ = build_issue_register([calculation, governance])
        priorities = {item["Issue_Code"]: item["Priority"] for item in register}
        self.assertEqual(priorities["ACTIVITY_MISSING"], "P1")
        self.assertEqual(priorities["CHEMISTRY_MISSING"], "P2")

    def test_top_contributor_uses_numeric_sort_and_safe_cumulative_share(self) -> None:
        rows = [
            _row("2024-A", emission="9", activity="1"),
            _row("2024-B", emission="100", activity="10"),
            _row("2024-C", emission="20", activity="2"),
        ]
        contributors, concentration = build_top_emission_contributors(rows, "2024")
        self.assertEqual(contributors[0]["Record_ID"], "2024-B")
        self.assertEqual(contributors[-1]["Cumulative_Share"], "1")
        self.assertLessEqual(Decimal(concentration["Top20_Emission_Share"]), 1)

    def test_factor_impact_is_ranked_without_recalculation(self) -> None:
        rows = [
            {
                "Year": "2024",
                "Record_ID": "2024-A",
                "Validated_Activity_kg": "10",
                "EF_2024": "2",
                "EF_2025": "1",
                "Factor_Impact_tCO2e": "-0.01",
                "Simulation_Flag": "TRUE",
                "Production_Eligible": "FALSE",
            },
            {
                "Year": "2024",
                "Record_ID": "2024-B",
                "Validated_Activity_kg": "20",
                "EF_2024": "2",
                "EF_2025": "1",
                "Factor_Impact_tCO2e": "-0.02",
                "Simulation_Flag": "TRUE",
                "Production_Eligible": "FALSE",
            },
        ]
        ranked = build_top_factor_impact(rows, "2024")
        self.assertEqual(ranked[0]["Record_ID"], "2024-B")
        self.assertEqual(ranked[0]["Factor_Impact_tCO2e"], "-0.02")
        self.assertEqual(ranked[0]["Current_EF"], "2")

    def test_lineage_summary_counts_complete_records(self) -> None:
        complete = _row("2024-A")
        incomplete = _row("2025-A")
        incomplete["Source_SHA256"] = ""
        summary = build_lineage_quality_summary(
            {"2024": [complete], "2025": [incomplete]}
        )
        self.assertEqual(summary["total"]["Complete_Lineage"], 1)
        self.assertEqual(summary["total"]["Incomplete_Lineage"], 1)


if __name__ == "__main__":
    unittest.main()
