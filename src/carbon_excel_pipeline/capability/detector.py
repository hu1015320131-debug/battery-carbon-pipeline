"""Capability-driven record and dataset decisions for WP6-2."""

from __future__ import annotations

import math
import re
from collections import Counter
from decimal import Decimal, InvalidOperation
from typing import Any

from .models import (
    ActivityPath,
    CapabilityStatus,
    PathDecision,
    RecordCapabilityResult,
    ValueStatus,
)


SUPPORTED_UNIT_WEIGHT_UNITS = frozenset({"g/PCS", "kg/PCS", "t/PCS"})
SUPPORTED_DIRECT_MASS_UNITS = frozenset({"g/year", "kg/year", "t/year"})
SUPPORTED_EF_UNITS = frozenset({"kgCO2e/kg", "tCO2e/t"})
SUPPORTED_HISTORICAL_GHG_UNITS = frozenset({"kgCO2e/year", "tCO2e/year"})

ANALYSIS_FIELDS = {
    "BUSINESS_UNIT_AVAILABLE": "Business_Unit",
    "CHEMISTRY_AVAILABLE": "Chemistry",
    "SUPPLIER_AVAILABLE": "Supplier",
    "PROJECT_AVAILABLE": "Project",
    "MODEL_AVAILABLE": "Model",
}
ANALYSIS_MISSING_CODES = {
    "Business_Unit": "BUSINESS_UNIT_MISSING",
    "Chemistry": "CHEMISTRY_MISSING",
    "Supplier": "SUPPLIER_MISSING",
    "Project": "PROJECT_MISSING",
    "Model": "MODEL_MISSING",
}
NUMERIC_PATTERN = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$")


def _present(value: Any) -> bool:
    return value is not None and (not isinstance(value, str) or bool(value.strip()))


def classify_numeric(value: Any) -> ValueStatus:
    if not _present(value):
        return ValueStatus.MISSING
    if isinstance(value, bool):
        return ValueStatus.NON_NUMERIC
    if isinstance(value, float) and not math.isfinite(value):
        return ValueStatus.NON_NUMERIC
    try:
        if isinstance(value, (int, float, Decimal)):
            number = Decimal(str(value))
        else:
            text = str(value).strip()
            if not NUMERIC_PATTERN.fullmatch(text):
                return ValueStatus.NON_NUMERIC
            number = Decimal(text)
    except (InvalidOperation, ValueError):
        return ValueStatus.NON_NUMERIC
    if number == 0:
        return ValueStatus.ZERO
    if number < 0:
        return ValueStatus.NEGATIVE
    return ValueStatus.VALID


def _numeric_reason(field_prefix: str, status: ValueStatus) -> str | None:
    return {
        ValueStatus.MISSING: f"{field_prefix}_VALUE_MISSING",
        ValueStatus.NON_NUMERIC: f"{field_prefix}_VALUE_NON_NUMERIC",
        ValueStatus.NEGATIVE: f"{field_prefix}_VALUE_NEGATIVE",
    }.get(status)


def _unit_reason(field_prefix: str, unit: str | None, supported: frozenset[str]) -> str | None:
    if not unit:
        return f"{field_prefix}_UNIT_MISSING"
    if unit not in supported:
        return f"{field_prefix}_UNIT_UNSUPPORTED"
    return None


def _path_reasons(
    *,
    values: dict[str, Any],
    units: dict[str, str | None],
    required_fields: tuple[tuple[str, str], ...],
    unit_field: str,
    unit_prefix: str,
    supported_units: frozenset[str],
) -> tuple[list[str], dict[str, ValueStatus]]:
    reasons: list[str] = []
    statuses: dict[str, ValueStatus] = {}
    for semantic_field, prefix in required_fields:
        if semantic_field not in values:
            reasons.append(f"{prefix}_FIELD_MISSING")
            statuses[semantic_field] = ValueStatus.MISSING
            continue
        status = classify_numeric(values.get(semantic_field))
        statuses[semantic_field] = status
        reason = _numeric_reason(prefix, status)
        if reason:
            reasons.append(reason)
    unit_reason = _unit_reason(unit_prefix, units.get(unit_field), supported_units)
    if unit_reason:
        reasons.append(unit_reason)
    return reasons, statuses


def detect_record_capabilities(
    record: dict[str, Any],
    *,
    units: dict[str, str | None],
) -> RecordCapabilityResult:
    values = dict(record.get("values") or {})
    source_row = int(record["Source_Row"])
    pcs_reasons, pcs_statuses = _path_reasons(
        values=values,
        units=units,
        required_fields=(
            ("Quantity_PCS", "PCS"),
            ("Unit_Weight", "UNIT_WEIGHT"),
        ),
        unit_field="Unit_Weight",
        unit_prefix="UNIT_WEIGHT",
        supported_units=SUPPORTED_UNIT_WEIGHT_UNITS,
    )
    direct_reasons, direct_statuses = _path_reasons(
        values=values,
        units=units,
        required_fields=(("Reported_Activity_Value", "REPORTED_ACTIVITY"),),
        unit_field="Reported_Activity_Value",
        unit_prefix="REPORTED_ACTIVITY",
        supported_units=SUPPORTED_DIRECT_MASS_UNITS,
    )
    supported: list[ActivityPath] = []
    if not pcs_reasons:
        supported.append(ActivityPath.PCS_WEIGHT_DERIVED)
    if not direct_reasons:
        supported.append(ActivityPath.DIRECT_REPORTED_MASS)
    path_decisions = [
        PathDecision(
            ActivityPath.PCS_WEIGHT_DERIVED,
            not pcs_reasons,
            tuple(pcs_reasons),
        ),
        PathDecision(
            ActivityPath.DIRECT_REPORTED_MASS,
            not direct_reasons,
            tuple(direct_reasons),
        ),
    ]

    ef_status = classify_numeric(values.get("EF_Value"))
    ef_reasons: list[str] = []
    if "EF_Value" not in values:
        ef_reasons.append("EF_VALUE_MISSING")
    else:
        ef_reason = _numeric_reason("EF", ef_status)
        if ef_reason:
            ef_reasons.append(ef_reason)
    ef_unit_reason = _unit_reason("EF", units.get("EF_Value"), SUPPORTED_EF_UNITS)
    if ef_unit_reason:
        ef_reasons.append(ef_unit_reason)
    factor_ready = not ef_reasons
    factor_source_available = _present(values.get("EF_Source"))

    historical_status = classify_numeric(values.get("Historical_GHG_Value"))
    historical_unit = units.get("Historical_GHG_Value")
    historical_result_available = (
        historical_status in {ValueStatus.VALID, ValueStatus.ZERO}
        and historical_unit in SUPPORTED_HISTORICAL_GHG_UNITS
    )
    activity_ready = bool(supported)
    emission_ready = activity_ready and factor_ready
    historical_validation_ready = emission_ready and historical_result_available

    analysis = {
        capability: _present(values.get(field_name))
        for capability, field_name in ANALYSIS_FIELDS.items()
    }
    warning_codes = [
        ANALYSIS_MISSING_CODES[field_name]
        for capability, field_name in ANALYSIS_FIELDS.items()
        if not analysis[capability]
    ]
    zero_warning_fields = {
        "Quantity_PCS": "PCS_VALUE_ZERO",
        "Unit_Weight": "UNIT_WEIGHT_VALUE_ZERO",
        "Reported_Activity_Value": "REPORTED_ACTIVITY_VALUE_ZERO",
        "EF_Value": "EF_VALUE_ZERO",
    }
    combined_statuses = {**pcs_statuses, **direct_statuses, "EF_Value": ef_status}
    warning_codes.extend(
        code
        for field_name, code in zero_warning_fields.items()
        if combined_statuses.get(field_name) == ValueStatus.ZERO
    )
    if factor_ready and not factor_source_available:
        warning_codes.append("EF_SOURCE_MISSING")
    if not historical_result_available:
        if historical_status == ValueStatus.MISSING:
            warning_codes.append("HISTORICAL_GHG_MISSING")
        elif historical_status == ValueStatus.NON_NUMERIC:
            warning_codes.append("HISTORICAL_GHG_VALUE_NON_NUMERIC")
        elif historical_status == ValueStatus.NEGATIVE:
            warning_codes.append("HISTORICAL_GHG_VALUE_NEGATIVE")
        if not historical_unit:
            warning_codes.append("HISTORICAL_GHG_UNIT_MISSING")
        elif historical_unit not in SUPPORTED_HISTORICAL_GHG_UNITS:
            warning_codes.append("HISTORICAL_GHG_UNIT_UNSUPPORTED")
    elif historical_status == ValueStatus.ZERO:
        warning_codes.append("HISTORICAL_GHG_VALUE_ZERO")

    blocking_codes: list[str] = []
    if not activity_ready:
        blocking_codes.extend(pcs_reasons)
        blocking_codes.extend(direct_reasons)
    if not factor_ready:
        blocking_codes.extend(ef_reasons)
    if not activity_ready:
        status = CapabilityStatus.INCAPABLE
    elif not factor_ready:
        status = CapabilityStatus.PARTIALLY_CAPABLE
    elif warning_codes:
        status = CapabilityStatus.CAPABLE_WITH_WARNING
    else:
        status = CapabilityStatus.CAPABLE

    value_statuses = combined_statuses
    value_statuses["Historical_GHG_Value"] = historical_status
    return RecordCapabilityResult(
        source_row=source_row,
        status=status,
        supported_activity_paths=supported,
        activity_ready=activity_ready,
        factor_ready=factor_ready,
        factor_source_available=factor_source_available,
        emission_ready=emission_ready,
        historical_result_available=historical_result_available,
        historical_validation_ready=historical_validation_ready,
        analysis_capabilities=analysis,
        value_statuses=value_statuses,
        path_decisions=path_decisions,
        warning_codes=sorted(set(warning_codes)),
        blocking_codes=sorted(set(blocking_codes)),
        formula_fields=sorted(record.get("formula_fields") or []),
    )


def _coverage(count: int, total: int) -> float:
    return count / total if total else 0.0


def detect_dataset_capabilities(payload: dict[str, Any]) -> dict[str, Any]:
    units = dict(payload.get("units") or {})
    results = [
        detect_record_capabilities(item, units=units)
        for item in payload.get("records", [])
    ]
    total = len(results)
    counts = {
        "activity_ready_count": sum(item.activity_ready for item in results),
        "factor_ready_count": sum(item.factor_ready for item in results),
        "emission_ready_count": sum(item.emission_ready for item in results),
        "historical_validation_ready_count": sum(
            item.historical_validation_ready for item in results
        ),
        "pcs_weight_derived_count": sum(
            ActivityPath.PCS_WEIGHT_DERIVED in item.supported_activity_paths
            for item in results
        ),
        "direct_reported_mass_count": sum(
            ActivityPath.DIRECT_REPORTED_MASS in item.supported_activity_paths
            for item in results
        ),
        "warning_count": sum(bool(item.warning_codes) for item in results),
        "incapable_count": sum(
            item.status == CapabilityStatus.INCAPABLE for item in results
        ),
    }
    status_counts = Counter(str(item.status) for item in results)
    warning_counts = Counter(code for item in results for code in item.warning_codes)
    blocking_counts = Counter(code for item in results for code in item.blocking_codes)
    if not results or counts["activity_ready_count"] == 0:
        dataset_status = CapabilityStatus.INCAPABLE
    elif counts["emission_ready_count"] < total:
        dataset_status = CapabilityStatus.PARTIALLY_CAPABLE
    elif counts["warning_count"]:
        dataset_status = CapabilityStatus.CAPABLE_WITH_WARNING
    else:
        dataset_status = CapabilityStatus.CAPABLE
    coverage = {
        key.removesuffix("_count") + "_coverage": _coverage(value, total)
        for key, value in counts.items()
        if key.endswith("_count") and key not in {"warning_count", "incapable_count"}
    }
    analysis_counts = {
        capability: sum(item.analysis_capabilities[capability] for item in results)
        for capability in ANALYSIS_FIELDS
    }
    dataset = {
        "schema_version": "WP6_2_DATASET_CAPABILITIES_V1",
        "status": dataset_status,
        "workbook_name": payload.get("workbook_name"),
        "input_fingerprint": payload.get("input_fingerprint"),
        "sheet_name": payload.get("sheet_name"),
        "sheet_index": payload.get("sheet_index"),
        "header_row": payload.get("header_row"),
        "denominator_definition": payload.get("denominator_definition"),
        "total_records": total,
        **counts,
        **coverage,
        "analysis_capability_counts": analysis_counts,
        "analysis_capability_coverage": {
            key: _coverage(value, total) for key, value in analysis_counts.items()
        },
        "record_status_counts": dict(sorted(status_counts.items())),
        "warning_code_counts": dict(sorted(warning_counts.items())),
        "blocking_code_counts": dict(sorted(blocking_counts.items())),
        "unit_metadata": units,
    }
    return {
        "dataset": dataset,
        "records": [item.to_dict() for item in results],
    }
