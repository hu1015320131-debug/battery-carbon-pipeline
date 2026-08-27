"""Live independent validation using only wp6_5.independent arithmetic."""

from __future__ import annotations

from typing import Any

from carbon_excel_pipeline.wp6_5.independent import (
    IndependentValidationInputError,
    calculate_direct_mass_activity,
    calculate_emission,
    calculate_pcs_weight_activity,
    compare_calculation,
    format_decimal,
)


def validate_live_canonical(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Recompute the current Run's rows. Does not read WP6-5 formal artifacts."""

    results: list[dict[str, Any]] = []
    for row in rows:
        method = str(row.get("Activity_Method", "")).strip()
        try:
            if method == "DIRECT_REPORTED_MASS":
                independent_activity = calculate_direct_mass_activity(
                    row.get("Original_Activity_Value"),
                    row.get("Original_Activity_Unit"),
                )
            elif method == "PCS_WEIGHT_DERIVED":
                independent_activity = calculate_pcs_weight_activity(
                    row.get("Quantity_PCS"),
                    row.get("Unit_Weight"),
                    row.get("Unit_Weight_Unit") or "g/PCS",
                )
            else:
                raise IndependentValidationInputError(f"unsupported method: {method}")
            independent_ef = row.get("EF_Value")
            independent_unit = row.get("EF_Unit") or "kgCO2e/kg"
            independent_emission = calculate_emission(
                independent_activity, independent_ef, independent_unit
            )
            comparison = compare_calculation(
                main_activity_kg=row.get("Activity_Data_kg"),
                independent_activity_kg=independent_activity,
                main_ef=row.get("EF_Value"),
                independent_ef=independent_ef,
                main_ef_unit=row.get("EF_Unit"),
                independent_ef_unit=independent_unit,
                main_emission_kgco2e=row.get("Emission_kgCO2e"),
                independent_emission_kgco2e=independent_emission,
                main_display_emission=row.get("Display_Emission_kgCO2e"),
                lineage_match=all(
                    str(row.get(field) or "").strip()
                    for field in ("Record_ID", "Source_SHA256", "Source_Sheet", "Source_Row")
                ),
                boundary_ready=str(row.get("Boundary_Ready", "")).upper() == "TRUE",
            )
            results.append(
                {
                    **row,
                    "Independent_Activity_kg": format_decimal(independent_activity),
                    "Independent_EF": str(independent_ef),
                    "Independent_Emission_kgCO2e": format_decimal(independent_emission),
                    **comparison,
                }
            )
        except IndependentValidationInputError as error:
            results.append(
                {
                    **row,
                    "Overall_Validation_Status": "INDEPENDENT_VALIDATION_BLOCKED",
                    "Reason_Codes": "VALIDATION_INPUT_MISSING",
                    "Validation_Error": str(error),
                }
            )
    return results
