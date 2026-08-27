from __future__ import annotations

import unittest

from carbon_excel_pipeline.capability.detector import (
    classify_numeric,
    detect_dataset_capabilities,
    detect_record_capabilities,
)
from carbon_excel_pipeline.capability.models import (
    ActivityPath,
    CapabilityStatus,
    ValueStatus,
)
from carbon_excel_pipeline.capability.policy import select_activity_path
from carbon_excel_pipeline.errors import PipelineUserError


BASE_UNITS = {
    "Unit_Weight": "g/PCS",
    "Reported_Activity_Value": "t/year",
    "EF_Value": "kgCO2e/kg",
    "Historical_GHG_Value": "tCO2e/year",
}


def record(row: int = 2, **values):
    return {"Source_Row": row, "values": values, "formula_fields": []}


class WP62CapabilityUnitTests(unittest.TestCase):
    def detect(self, values: dict, units: dict | None = None):
        return detect_record_capabilities(
            record(**values), units=BASE_UNITS if units is None else units
        )

    def test_pcs_and_unit_weight_support_derived_path(self) -> None:
        result = self.detect({"Quantity_PCS": 10, "Unit_Weight": 2})
        self.assertEqual(result.supported_activity_paths, [ActivityPath.PCS_WEIGHT_DERIVED])

    def test_direct_mass_with_t_per_year_supports_direct_path(self) -> None:
        result = self.detect({"Reported_Activity_Value": 1.2})
        self.assertEqual(result.supported_activity_paths, [ActivityPath.DIRECT_REPORTED_MASS])
        self.assertTrue(result.activity_ready)

    def test_both_paths_are_kept_without_detector_preference(self) -> None:
        result = self.detect(
            {"Quantity_PCS": 10, "Unit_Weight": 2, "Reported_Activity_Value": 0.02}
        )
        self.assertEqual(
            result.supported_activity_paths,
            [ActivityPath.PCS_WEIGHT_DERIVED, ActivityPath.DIRECT_REPORTED_MASS],
        )
        self.assertEqual(len(result.supported_activity_paths), 2)

    def test_processing_policy_explicitly_selects_supported_path(self) -> None:
        supported = [ActivityPath.PCS_WEIGHT_DERIVED, ActivityPath.DIRECT_REPORTED_MASS]
        self.assertEqual(
            select_activity_path(supported, requested_path=ActivityPath.DIRECT_REPORTED_MASS),
            ActivityPath.DIRECT_REPORTED_MASS,
        )

    def test_processing_policy_rejects_unsupported_path(self) -> None:
        with self.assertRaises(PipelineUserError):
            select_activity_path(
                [ActivityPath.DIRECT_REPORTED_MASS],
                requested_path=ActivityPath.PCS_WEIGHT_DERIVED,
            )

    def test_activity_and_factor_make_emission_ready(self) -> None:
        result = self.detect(
            {"Reported_Activity_Value": 1.2, "EF_Value": 14, "EF_Source": "历史因子"}
        )
        self.assertTrue(result.emission_ready)

    def test_missing_pcs_does_not_block_direct_path(self) -> None:
        result = self.detect({"Reported_Activity_Value": 1.2})
        self.assertTrue(result.activity_ready)
        self.assertNotIn("PCS_FIELD_MISSING", result.blocking_codes)
        pcs = next(item for item in result.path_decisions if item.path == ActivityPath.PCS_WEIGHT_DERIVED)
        self.assertIn("PCS_FIELD_MISSING", pcs.reason_codes)

    def test_missing_unit_weight_does_not_block_direct_path(self) -> None:
        result = self.detect({"Quantity_PCS": 10, "Reported_Activity_Value": 1.2})
        self.assertEqual(result.supported_activity_paths, [ActivityPath.DIRECT_REPORTED_MASS])

    def test_direct_mass_unit_is_not_guessed(self) -> None:
        units = {**BASE_UNITS, "Reported_Activity_Value": None}
        result = self.detect({"Reported_Activity_Value": 1.2}, units)
        self.assertFalse(result.activity_ready)
        self.assertIn("REPORTED_ACTIVITY_UNIT_MISSING", result.blocking_codes)

    def test_non_numeric_activity_is_not_ready(self) -> None:
        result = self.detect({"Reported_Activity_Value": "unknown"})
        self.assertFalse(result.activity_ready)
        self.assertIn("REPORTED_ACTIVITY_VALUE_NON_NUMERIC", result.blocking_codes)

    def test_missing_ef_keeps_activity_but_not_emission_ready(self) -> None:
        result = self.detect({"Reported_Activity_Value": 1.2})
        self.assertTrue(result.activity_ready)
        self.assertFalse(result.emission_ready)
        self.assertEqual(result.status, CapabilityStatus.PARTIALLY_CAPABLE)

    def test_missing_chemistry_does_not_block_emission(self) -> None:
        result = self.detect(
            {"Reported_Activity_Value": 1.2, "EF_Value": 14, "EF_Source": "历史因子"}
        )
        self.assertTrue(result.emission_ready)
        self.assertIn("CHEMISTRY_MISSING", result.warning_codes)

    def test_all_activity_paths_failed_is_incapable(self) -> None:
        result = self.detect({"EF_Value": 14})
        self.assertFalse(result.activity_ready)
        self.assertEqual(result.status, CapabilityStatus.INCAPABLE)

    def test_partial_dataset_does_not_fail_all_records(self) -> None:
        payload = {
            "units": BASE_UNITS,
            "records": [
                record(2, Reported_Activity_Value=1.2, EF_Value=14),
                record(3, Reported_Activity_Value="bad", EF_Value=14),
            ],
        }
        result = detect_dataset_capabilities(payload)["dataset"]
        self.assertEqual(result["total_records"], 2)
        self.assertEqual(result["activity_ready_count"], 1)
        self.assertEqual(result["status"], CapabilityStatus.PARTIALLY_CAPABLE)

    def test_zero_negative_and_non_numeric_are_distinct(self) -> None:
        self.assertEqual(classify_numeric(0), ValueStatus.ZERO)
        self.assertEqual(classify_numeric(-1), ValueStatus.NEGATIVE)
        self.assertEqual(classify_numeric("x"), ValueStatus.NON_NUMERIC)

    def test_zero_is_supported_with_warning_for_downstream_decision(self) -> None:
        result = self.detect({"Reported_Activity_Value": 0, "EF_Value": 0})
        self.assertTrue(result.activity_ready)
        self.assertTrue(result.factor_ready)
        self.assertTrue(result.emission_ready)
        self.assertIn("REPORTED_ACTIVITY_VALUE_ZERO", result.warning_codes)
        self.assertIn("EF_VALUE_ZERO", result.warning_codes)


if __name__ == "__main__":
    unittest.main()
