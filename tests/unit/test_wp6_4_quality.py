from __future__ import annotations

import unittest

from carbon_excel_pipeline.wp6_4.quality import (
    assess_quality_layers,
    ef_audit,
    forward_fill_audit,
    preferred_precision_status,
)


class WP64QualityTests(unittest.TestCase):
    @staticmethod
    def _result() -> dict[str, str]:
        return {
            "Activity_Data_Normalized_kg": "10",
            "Activity_Unit_Normalized": "kg/year",
            "EF_Value_Normalized_kgCO2e_per_kg": "2",
            "EF_Unit_Normalized": "kgCO2e/kg",
            "Raw_Emission_kgCO2e": "20",
        }

    def test_calculation_and_governance_qc_are_separate(self) -> None:
        quality = assess_quality_layers(
            activity={"QC_Status": "WARNING", "Issue_Code": "CHEMISTRY_UNKNOWN"},
            result=self._result(),
            standard={
                "Business_Unit": "二部",
                "Activity_Category": "电芯",
                "Product_Description": "SYNA cell",
                "Chemistry": "UNKNOWN",
                "Supplier_Name": "Synthetic Supplier A",
                "Project_Code": "P1",
                "Cell_Model": "M1",
            },
        )
        self.assertEqual(quality["Calculation_QC"], "PASS")
        self.assertEqual(quality["Governance_QC"], "WARNING")
        self.assertEqual(quality["Overall_Status"], "PASS_WITH_WARNING")
        self.assertTrue(quality["Emission_Ready"])

    def test_boundary_ready_is_independent_from_emission_ready(self) -> None:
        quality = assess_quality_layers(
            activity={"QC_Status": "PASS", "Issue_Code": "NONE"},
            result=self._result(),
            standard={
                "Business_Unit": "",
                "Activity_Category": "电芯",
                "Product_Description": "SYNA cell",
                "Chemistry": "LCO",
                "Supplier_Name": "Synthetic Supplier A",
                "Project_Code": "P1",
                "Cell_Model": "M1",
            },
        )
        self.assertTrue(quality["Emission_Ready"])
        self.assertFalse(quality["Boundary_Ready"])
        self.assertEqual(quality["Overall_Status"], "BLOCKED")

    def test_ef_audit_reports_unique_min_max_and_weighted_value(self) -> None:
        audit = ef_audit(
            [
                {"Activity_Data_Normalized_kg": "1", "EF_Value_Normalized_kgCO2e_per_kg": "2"},
                {"Activity_Data_Normalized_kg": "3", "EF_Value_Normalized_kgCO2e_per_kg": "4"},
            ]
        )
        self.assertEqual(audit["EF_Unique_Count"], 2)
        self.assertEqual(audit["EF_Min"], "2")
        self.assertEqual(audit["EF_Max"], "4")
        self.assertEqual(audit["Activity_Weighted_EF"], "3.5")

    def test_forward_fill_summary_preserves_source_categories(self) -> None:
        audit = forward_fill_audit(
            [
                {
                    "Business_Unit_Source": "ORIGINAL",
                    "Purchase_Type_Source": "CONTROLLED_FORWARD_FILL",
                    "Purchase_Category_Source": "MISSING",
                },
                {
                    "Business_Unit_Source": "CONTROLLED_FORWARD_FILL",
                    "Purchase_Type_Source": "ORIGINAL",
                    "Purchase_Category_Source": "ORIGINAL",
                },
            ]
        )
        self.assertEqual(audit["Business_Unit"]["Original_Count"], 1)
        self.assertEqual(audit["Business_Unit"]["Forward_Filled_Count"], 1)
        self.assertEqual(audit["Purchase_Category"]["Missing_Count"], 1)

    def test_formula_cache_precision_term_replaces_rounding_term(self) -> None:
        self.assertEqual(
            preferred_precision_status("PASS_WITH_REPORTED_ROUNDING_DIFFERENCE"),
            "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE",
        )


if __name__ == "__main__":
    unittest.main()
