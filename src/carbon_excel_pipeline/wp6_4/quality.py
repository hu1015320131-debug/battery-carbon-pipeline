"""Shared WP6 quality layers and audit summaries."""

from __future__ import annotations

from collections import Counter
from decimal import Decimal, InvalidOperation
from typing import Any, Iterable


GOVERNANCE_FIELDS = ("Chemistry", "Supplier_Name", "Project_Code", "Cell_Model")
BOUNDARY_FIELDS = ("Business_Unit", "Activity_Category", "Product_Description")


def _present(value: Any) -> bool:
    return value is not None and str(value).strip() not in {"", "UNKNOWN"}


def _positive_decimal(value: Any) -> bool:
    try:
        parsed = Decimal(str(value).strip())
    except (InvalidOperation, ValueError, AttributeError):
        return False
    return parsed.is_finite() and parsed > 0


def assess_quality_layers(
    *,
    activity: dict[str, Any],
    result: dict[str, Any],
    standard: dict[str, Any],
) -> dict[str, Any]:
    """Keep mathematical readiness separate from governance completeness."""

    calculation_issues: list[str] = []
    if not _positive_decimal(result.get("Activity_Data_Normalized_kg")):
        calculation_issues.append("ACTIVITY_INVALID")
    if result.get("Activity_Unit_Normalized") != "kg/year":
        calculation_issues.append("ACTIVITY_UNIT_INVALID")
    if not _positive_decimal(result.get("EF_Value_Normalized_kgCO2e_per_kg")):
        calculation_issues.append("EF_INVALID")
    if result.get("EF_Unit_Normalized") != "kgCO2e/kg":
        calculation_issues.append("EF_UNIT_INVALID")
    if not _positive_decimal(result.get("Raw_Emission_kgCO2e")):
        calculation_issues.append("EMISSION_INVALID")

    governance_issues = [
        f"{field.upper()}_MISSING"
        for field in GOVERNANCE_FIELDS
        if not _present(standard.get(field))
    ]
    legacy_governance = str(activity.get("QC_Status", "")).strip()
    legacy_issue = str(activity.get("Issue_Code", "")).strip()
    if legacy_governance == "WARNING" and legacy_issue not in {"", "NONE"}:
        governance_issues.extend(
            f"LEGACY_{code}" for code in legacy_issue.split(";") if code
        )

    boundary_issues = [
        f"{field.upper()}_MISSING"
        for field in BOUNDARY_FIELDS
        if not _present(standard.get(field))
    ]
    activity_ready = _positive_decimal(result.get("Activity_Data_Normalized_kg")) and (
        result.get("Activity_Unit_Normalized") == "kg/year"
    )
    factor_ready = _positive_decimal(
        result.get("EF_Value_Normalized_kgCO2e_per_kg")
    ) and result.get("EF_Unit_Normalized") == "kgCO2e/kg"
    emission_ready = activity_ready and factor_ready and _positive_decimal(
        result.get("Raw_Emission_kgCO2e")
    )
    boundary_ready = not boundary_issues
    calculation_qc = "PASS" if not calculation_issues else "BLOCKED"
    governance_qc = "WARNING" if governance_issues else "PASS"
    if calculation_qc == "BLOCKED" or not boundary_ready:
        overall = "BLOCKED"
    elif governance_qc == "WARNING":
        overall = "PASS_WITH_WARNING"
    else:
        overall = "PASS"
    return {
        "Activity_Ready": activity_ready,
        "Factor_Ready": factor_ready,
        "Emission_Ready": emission_ready,
        "Boundary_Ready": boundary_ready,
        "Calculation_QC": calculation_qc,
        "Governance_QC": governance_qc,
        "Overall_Status": overall,
        "Calculation_Issue_Codes": "|".join(sorted(set(calculation_issues))),
        "Governance_Issue_Codes": "|".join(sorted(set(governance_issues))),
        "Boundary_Issue_Codes": "|".join(sorted(set(boundary_issues))),
    }


def ef_audit(rows: Iterable[dict[str, Any]]) -> dict[str, str | int]:
    pairs = [
        (
            Decimal(str(row["Activity_Data_Normalized_kg"])),
            Decimal(str(row["EF_Value_Normalized_kgCO2e_per_kg"])),
        )
        for row in rows
    ]
    if not pairs:
        return {
            "EF_Unique_Count": 0,
            "EF_Min": "",
            "EF_Max": "",
            "Activity_Weighted_EF": "",
        }
    factors = [factor for _, factor in pairs]
    activity_total = sum((activity for activity, _ in pairs), Decimal("0"))
    weighted = (
        sum((activity * factor for activity, factor in pairs), Decimal("0"))
        / activity_total
        if activity_total
        else Decimal("0")
    )
    return {
        "EF_Unique_Count": len(set(factors)),
        "EF_Min": format(min(factors), "f"),
        "EF_Max": format(max(factors), "f"),
        "Activity_Weighted_EF": format(weighted, "f"),
    }


def forward_fill_audit(rows: Iterable[dict[str, Any]]) -> dict[str, Any]:
    fields = ("Business_Unit", "Purchase_Type", "Purchase_Category")
    result: dict[str, Any] = {}
    for field in fields:
        counts = Counter(str(row.get(f"{field}_Source", "MISSING")) for row in rows)
        result[field] = {
            "Original_Count": counts.get("ORIGINAL", 0),
            "Forward_Filled_Count": counts.get("CONTROLLED_FORWARD_FILL", 0),
            "Missing_Count": counts.get("MISSING", 0),
        }
    return result


def preferred_precision_status(value: str) -> str:
    mapping = {
        "PASS_WITH_REPORTED_ROUNDING_DIFFERENCE": (
            "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE"
        ),
        "PASS_WITH_REPORTED_FORMULA_CACHE_PRECISION_DIFFERENCE": (
            "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE"
        ),
    }
    return mapping.get(value, value)
