"""Pure transformations for WP6-3 boundary, adapters, QC, and validation."""

from __future__ import annotations

import re
from copy import deepcopy
from decimal import Decimal, InvalidOperation
from typing import Any

from carbon_excel_pipeline.calculation.day7_calculator import (
    canonical_decimal,
    multiply_emission_decimal,
)


CONTROLLED_CONTEXT_FIELDS = (
    "Business_Unit",
    "Purchase_Type",
    "Purchase_Category",
)
FORBIDDEN_FORWARD_FILL_FIELDS = (
    "Product_Description",
    "Reported_Activity_Value",
    "EF_Value",
    "EF_Source",
    "Historical_GHG_Value",
)
NON_CRITICAL_ANALYSIS_FIELDS = ("Chemistry", "Supplier", "Project", "Model")
MASS_FACTORS_TO_KG = {
    "g/year": Decimal("0.001"),
    "kg/year": Decimal("1"),
    "t/year": Decimal("1000"),
}
EF_FACTORS_TO_KGCO2E_PER_KG = {
    "kgCO2e/kg": Decimal("1"),
    "tCO2e/t": Decimal("1"),
}
HISTORICAL_FACTORS_TO_TCO2E = {
    "kgCO2e/year": Decimal("0.001"),
    "tCO2e/year": Decimal("1"),
}


def is_present(value: Any) -> bool:
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


def parse_decimal(value: Any) -> tuple[Decimal | None, str]:
    if not is_present(value):
        return None, "MISSING"
    if isinstance(value, bool):
        return None, "NON_NUMERIC"
    try:
        parsed = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        return None, "NON_NUMERIC"
    if not parsed.is_finite():
        return None, "NON_NUMERIC"
    if parsed < 0:
        return parsed, "NEGATIVE"
    if parsed == 0:
        return parsed, "ZERO"
    return parsed, "VALID"


def controlled_forward_fill(
    records: list[dict[str, Any]],
    *,
    fields: tuple[str, ...] = CONTROLLED_CONTEXT_FIELDS,
) -> list[dict[str, Any]]:
    """Fill only approved context fields while preserving physical row lineage."""

    if not set(fields).issubset(CONTROLLED_CONTEXT_FIELDS):
        raise ValueError("forward-fill fields must be approved context fields")
    ordered = sorted(records, key=lambda item: int(item["Source_Row"]))
    source_rows = [int(item["Source_Row"]) for item in ordered]
    if len(source_rows) != len(set(source_rows)) or any(row <= 0 for row in source_rows):
        raise ValueError("Source_Row must contain unique positive physical row numbers")
    last_seen: dict[str, Any] = {}
    result: list[dict[str, Any]] = []
    for record in ordered:
        output = deepcopy(record)
        values = dict(output.get("values") or {})
        original_values: dict[str, Any] = {}
        value_sources: dict[str, str] = {}
        for field in fields:
            original = values.get(field)
            original_values[field] = original
            if is_present(original):
                last_seen[field] = original
                value_sources[field] = "ORIGINAL"
            elif field in last_seen:
                values[field] = last_seen[field]
                value_sources[field] = "CONTROLLED_FORWARD_FILL"
            else:
                value_sources[field] = "MISSING"
        output["values"] = values
        output["original_context_values"] = original_values
        output["context_value_sources"] = value_sources
        result.append(output)
    return result


def filter_2024_boundary(
    records: list[dict[str, Any]],
    *,
    business_unit: str,
    category_root: str,
    product_marker: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Apply three explicit business filters; counts are evidence, never inputs."""

    normalized_unit = business_unit.strip().casefold()
    normalized_category = category_root.strip().casefold()
    marker = re.compile(re.escape(product_marker.strip()), re.IGNORECASE)
    unit_rows = [
        item
        for item in records
        if str(item.get("values", {}).get("Business_Unit") or "").strip().casefold()
        == normalized_unit
    ]
    cell_rows = []
    for item in unit_rows:
        category = str(item.get("values", {}).get("Purchase_Category") or "").strip()
        root = category.split(".", 1)[0].strip().casefold()
        if root == normalized_category:
            cell_rows.append(item)
    marker_rows = [
        item
        for item in cell_rows
        if marker.search(str(item.get("values", {}).get("Product_Description") or ""))
    ]
    audit = [
        {
            "Stage": "BUSINESS_UNIT",
            "Input_Count": len(records),
            "Selected_Count": len(unit_rows),
            "Excluded_Count": len(records) - len(unit_rows),
            "Filter_Rule": f"Business_Unit exact = {business_unit}",
        },
        {
            "Stage": "PURCHASE_CATEGORY",
            "Input_Count": len(unit_rows),
            "Selected_Count": len(cell_rows),
            "Excluded_Count": len(unit_rows) - len(cell_rows),
            "Filter_Rule": f"Purchase_Category root exact = {category_root}",
        },
        {
            "Stage": "PRODUCT_DESCRIPTION",
            "Input_Count": len(cell_rows),
            "Selected_Count": len(marker_rows),
            "Excluded_Count": len(cell_rows) - len(marker_rows),
            "Filter_Rule": f"Product_Description contains marker {product_marker} (case-insensitive)",
        },
    ]
    return marker_rows, audit


def assign_2024_record_ids(
    records: list[dict[str, Any]], *, year: int = 2024
) -> list[dict[str, Any]]:
    ordered = sorted(records, key=lambda item: int(item["Source_Row"]))
    output: list[dict[str, Any]] = []
    for index, record in enumerate(ordered, start=1):
        item = deepcopy(record)
        item["Record_ID"] = f"{year}-DY2-SYNA-DX{index:06d}"
        output.append(item)
    return output


def adapt_direct_mass(value: Any, unit: str | None) -> dict[str, Any]:
    parsed, status = parse_decimal(value)
    warnings: list[str] = []
    blocking: list[str] = []
    if status == "ZERO":
        warnings.append("ZERO_ACTIVITY")
    elif status != "VALID":
        blocking.append(f"ACTIVITY_{status}")
    factor = MASS_FACTORS_TO_KG.get(unit or "")
    if not unit:
        blocking.append("ACTIVITY_UNIT_MISSING")
    elif factor is None:
        blocking.append("ACTIVITY_UNIT_UNSUPPORTED")
    activity_kg = parsed * factor if parsed is not None and factor is not None and not blocking else None
    return {
        "original_value": value,
        "original_unit": unit,
        "conversion_factor": factor,
        "activity_kg": activity_kg,
        "warnings": warnings,
        "blocking": sorted(set(blocking)),
    }


def adapt_historical_ef(value: Any, unit: str | None, source: Any) -> dict[str, Any]:
    parsed, status = parse_decimal(value)
    warnings: list[str] = []
    blocking: list[str] = []
    if status == "ZERO":
        warnings.append("ZERO_EF")
    elif status != "VALID":
        blocking.append(f"EF_{status}")
    factor = EF_FACTORS_TO_KGCO2E_PER_KG.get(unit or "")
    if not unit:
        blocking.append("EF_UNIT_MISSING")
    elif factor is None:
        blocking.append("EF_UNIT_UNSUPPORTED")
    if not is_present(source):
        warnings.append("EF_SOURCE_MISSING")
    normalized = parsed * factor if parsed is not None and factor is not None and not blocking else None
    return {
        "original_value": value,
        "original_unit": unit,
        "source": source,
        "normalized_value": normalized,
        "warnings": warnings,
        "blocking": sorted(set(blocking)),
    }


def adapt_historical_ghg(value: Any, unit: str | None) -> dict[str, Any]:
    parsed, status = parse_decimal(value)
    warnings: list[str] = []
    if status not in {"VALID", "ZERO"}:
        warnings.append(f"HISTORICAL_GHG_{status}")
    factor = HISTORICAL_FACTORS_TO_TCO2E.get(unit or "")
    if not unit:
        warnings.append("HISTORICAL_GHG_UNIT_MISSING")
    elif factor is None:
        warnings.append("HISTORICAL_GHG_UNIT_UNSUPPORTED")
    normalized = parsed * factor if parsed is not None and factor is not None and not warnings else None
    return {
        "original_value": value,
        "original_unit": unit,
        "normalized_tco2e": normalized,
        "warnings": warnings,
    }


def calculate_and_validate(
    *,
    activity_kg: Decimal,
    ef_value: Decimal,
    historical_tco2e: Decimal | None,
    difference_threshold_tco2e: Decimal,
) -> dict[str, Any]:
    emission_kg, emission_t = multiply_emission_decimal(activity_kg, ef_value)
    if historical_tco2e is None:
        return {
            "emission_kg": emission_kg,
            "emission_t": emission_t,
            "difference_t": None,
            "difference_percent": None,
            "validation_status": "HISTORICAL_NOT_AVAILABLE",
        }
    difference = emission_t - historical_tco2e
    difference_percent = (
        difference / historical_tco2e * Decimal("100")
        if historical_tco2e != 0
        else None
    )
    if difference == 0:
        status = "PASS_EXACT"
    elif abs(difference) <= difference_threshold_tco2e:
        status = "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE"
    else:
        status = "DIFFERENCE_REQUIRES_REVIEW"
    return {
        "emission_kg": emission_kg,
        "emission_t": emission_t,
        "difference_t": difference,
        "difference_percent": difference_percent,
        "validation_status": status,
    }


def decimal_text(value: Decimal | None) -> str:
    return "" if value is None else canonical_decimal(value)
