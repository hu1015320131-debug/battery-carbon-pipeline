"""Validation-only arithmetic for WP6-5.

This module intentionally uses only Python's standard-library Decimal engine.
It must never import a production activity adapter, calculation function, or
pipeline module.
"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from typing import Any


DISPLAY_QUANTUM = Decimal("0.000001")
DIRECT_MASS_FACTORS = {
    "t/year": Decimal("1000"),
    "kg/year": Decimal("1"),
    "g/year": Decimal("0.001"),
}


class IndependentValidationInputError(ValueError):
    """Raised when a validation input cannot be interpreted deterministically."""


def decimal_value(value: Any, field: str) -> Decimal:
    """Parse a finite Decimal without passing through binary floating point."""

    if isinstance(value, bool) or value is None or str(value).strip() == "":
        raise IndependentValidationInputError(f"{field} is missing")
    try:
        parsed = Decimal(str(value).strip())
    except (InvalidOperation, ValueError) as error:
        raise IndependentValidationInputError(f"{field} is not decimal: {value}") from error
    if not parsed.is_finite():
        raise IndependentValidationInputError(f"{field} is not finite: {value}")
    return parsed


def normalize_mass_unit(value: Any) -> str:
    text = str(value or "").strip().casefold().replace(" ", "")
    aliases = {
        "t/year": "t/year",
        "t/年": "t/year",
        "kg/year": "kg/year",
        "kg/年": "kg/year",
        "g/year": "g/year",
        "g/年": "g/year",
    }
    if text not in aliases:
        raise IndependentValidationInputError(f"unsupported activity unit: {value}")
    return aliases[text]


def normalize_unit_weight_unit(value: Any) -> str:
    text = str(value or "").strip().casefold().replace(" ", "")
    aliases = {"g/pcs": "g/PCS", "g/pc": "g/PCS", "g/个": "g/PCS"}
    if text not in aliases:
        raise IndependentValidationInputError(f"unsupported unit-weight unit: {value}")
    return aliases[text]


def normalize_ef_unit(value: Any) -> str:
    text = (
        str(value or "")
        .strip()
        .casefold()
        .replace(" ", "")
        .replace("₂", "2")
    )
    aliases = {
        "kgco2e/kg": "kgCO2e/kg",
        "kgco2/kg": "kgCO2e/kg",
    }
    if text not in aliases:
        raise IndependentValidationInputError(f"unsupported EF unit: {value}")
    return aliases[text]


def calculate_direct_mass_activity(value: Any, unit: Any) -> Decimal:
    """Independently convert a directly reported annual mass to kg/year."""

    activity = decimal_value(value, "Original_Activity_Value")
    normalized_unit = normalize_mass_unit(unit)
    with localcontext() as context:
        context.prec = 50
        return activity * DIRECT_MASS_FACTORS[normalized_unit]


def calculate_pcs_weight_activity(
    quantity_pcs: Any,
    unit_weight: Any,
    unit_weight_unit: Any = "g/PCS",
) -> Decimal:
    """Independently calculate PCS × g/PCS ÷ 1000 as kg/year."""

    quantity = decimal_value(quantity_pcs, "Quantity_PCS")
    weight = decimal_value(unit_weight, "Unit_Weight")
    normalize_unit_weight_unit(unit_weight_unit)
    with localcontext() as context:
        context.prec = 50
        return quantity * weight / Decimal("1000")


def calculate_emission(activity_kg: Any, ef_value: Any, ef_unit: Any) -> Decimal:
    """Independently calculate kg activity × kgCO2e/kg."""

    activity = decimal_value(activity_kg, "Independent_Activity_kg")
    factor = decimal_value(ef_value, "Independent_EF")
    normalize_ef_unit(ef_unit)
    with localcontext() as context:
        context.prec = 50
        return activity * factor


def display_six(value: Any) -> Decimal:
    parsed = decimal_value(value, "display value")
    return parsed.quantize(DISPLAY_QUANTUM, rounding=ROUND_HALF_UP)


def format_decimal(value: Decimal) -> str:
    return format(value, "f")


def compare_calculation(
    *,
    main_activity_kg: Any,
    independent_activity_kg: Any,
    main_ef: Any,
    independent_ef: Any,
    main_ef_unit: Any,
    independent_ef_unit: Any,
    main_emission_kgco2e: Any,
    independent_emission_kgco2e: Any,
    main_display_emission: Any | None = None,
    lineage_match: bool = True,
    boundary_ready: bool = True,
    activity_input_match: bool = True,
) -> dict[str, str]:
    """Compare raw, display, unit, boundary, and lineage facts exactly."""

    main_activity = decimal_value(main_activity_kg, "Main_Activity_kg")
    independent_activity = decimal_value(
        independent_activity_kg, "Independent_Activity_kg"
    )
    main_factor = decimal_value(main_ef, "Main_EF")
    independent_factor = decimal_value(independent_ef, "Independent_EF")
    main_emission = decimal_value(main_emission_kgco2e, "Main_Emission_kgCO2e")
    independent_emission = decimal_value(
        independent_emission_kgco2e, "Independent_Emission_kgCO2e"
    )
    activity_difference = main_activity - independent_activity
    ef_difference = main_factor - independent_factor
    emission_difference = main_emission - independent_emission
    try:
        main_unit = normalize_ef_unit(main_ef_unit)
    except IndependentValidationInputError:
        main_unit = f"UNSUPPORTED:{main_ef_unit}"
    try:
        independent_unit = normalize_ef_unit(independent_ef_unit)
    except IndependentValidationInputError:
        independent_unit = f"UNSUPPORTED:{independent_ef_unit}"
    main_display = display_six(
        main_display_emission if main_display_emission not in {None, ""} else main_emission
    )
    independent_display = display_six(independent_emission)
    display_difference = main_display - independent_display

    reasons: list[str] = []
    if not activity_input_match:
        reasons.append("ACTIVITY_INPUT_MISMATCH")
    if activity_difference != 0:
        reasons.append("ACTIVITY_CONVERSION_MISMATCH")
    if main_unit != independent_unit:
        reasons.append("EF_UNIT_MISMATCH")
    if ef_difference != 0:
        reasons.append("EF_VALUE_MISMATCH")
    if emission_difference != 0:
        if display_difference == 0:
            reasons.append("ROUNDING_ONLY_DIFFERENCE")
        else:
            reasons.append("EMISSION_CALCULATION_MISMATCH")
    if not lineage_match:
        reasons.append("LINEAGE_MISMATCH")
    if not boundary_ready:
        reasons.append("VALIDATION_SCOPE_ANOMALY")

    exact = not reasons
    return {
        "Activity_Difference": format_decimal(activity_difference),
        "Activity_Validation_Status": "PASS_EXACT" if activity_difference == 0 else "FAIL",
        "EF_Difference": format_decimal(ef_difference),
        "EF_Validation_Status": (
            "PASS_EXACT" if ef_difference == 0 and main_unit == independent_unit else "FAIL"
        ),
        "Emission_Difference": format_decimal(emission_difference),
        "Emission_Validation_Status": "PASS_EXACT" if emission_difference == 0 else "FAIL",
        "Main_Emission_Display_6dp": format_decimal(main_display),
        "Independent_Emission_Display_6dp": format_decimal(independent_display),
        "Display_Difference": format_decimal(display_difference),
        "Display_Validation_Status": "PASS" if display_difference == 0 else "FAIL",
        "Lineage_Validation_Status": "PASS" if lineage_match else "FAIL",
        "Boundary_Validation_Status": "PASS" if boundary_ready else "FAIL",
        "Reason_Codes": "|".join(reasons),
        "Overall_Validation_Status": (
            "INDEPENDENT_CALCULATION_PASS" if exact else "INDEPENDENT_VALIDATION_FAIL"
        ),
    }
