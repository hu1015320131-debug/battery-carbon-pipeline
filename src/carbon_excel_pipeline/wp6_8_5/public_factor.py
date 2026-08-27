"""Apply ledger-backed 2025 cell-category historical simulation factor."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any

from carbon_excel_pipeline.wp6_8_5.cell_scope import is_cell_category


POLICY_ID = "SYNTHETIC_CELL_CATEGORY_FACTOR_V1"
BOUNDARY_ID = "SYNTHETIC_CELL_CATEGORY_SCOPE_V1"
ALLOWED_UNITS = {"合成一部", "合成二部", "合成一部", "合成二部"}


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _decimal(value: Any) -> Decimal | None:
    try:
        return Decimal(str(value)) if value not in (None, "") else None
    except (InvalidOperation, ValueError):
        return None


def _without_codes(value: Any, removed: set[str]) -> str:
    codes = [code for code in _text(value).replace(";", "|").split("|") if code]
    return "|".join(code for code in codes if code not in removed)


def apply_public_cell_factor(
    rows: list[dict[str, Any]], evidence: dict[str, Any] | None
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    if not evidence or evidence.get("status") != "PASS":
        return [dict(row) for row in rows], {
            "status": "NOT_APPLIED",
            "policy_id": POLICY_ID,
            "reason": "VALID_LEDGER_EVIDENCE_NOT_AVAILABLE",
            "applied_records": 0,
        }
    factor = _decimal(evidence.get("EF_Value"))
    if factor is None or factor <= 0:
        return [dict(row) for row in rows], {
            "status": "NOT_APPLIED",
            "policy_id": POLICY_ID,
            "reason": "LEDGER_CELL_FACTOR_INVALID",
            "applied_records": 0,
        }
    source = (
        f"{evidence.get('Source_File')}/{evidence.get('Factor_Source_Sheet')}"
        f"#{evidence.get('Factor_Source_Cell')}"
    )
    output: list[dict[str, Any]] = []
    applied = 0
    skipped_non_cell = 0
    skipped_scope = 0
    skipped_activity = 0
    for row in rows:
        item = dict(row)
        if not is_cell_category(item.get("Purchase_Category")):
            skipped_non_cell += 1
            output.append(item)
            continue
        if _text(item.get("Business_Unit")) not in ALLOWED_UNITS:
            skipped_scope += 1
            output.append(item)
            continue
        activity = _decimal(item.get("Activity_Data_kg"))
        if activity is None or activity <= 0:
            skipped_activity += 1
            output.append(item)
            continue
        emission_kg = activity * factor
        blocking = _without_codes(
            item.get("Blocking_Codes"),
            {"FACTOR_NOT_AVAILABLE", "BOUNDARY_POLICY_NOT_AVAILABLE"},
        )
        governance_warning = any(
            not _text(item.get(field))
            for field in ("Chemistry", "Supplier", "Project", "Model", "Customer")
        )
        item.update(
            {
                "EF_Value": format(factor, "f"),
                "EF_Unit": "kgCO2e/kg",
                "EF_Source": source,
                "EF_Usage": "HISTORICAL_SIMULATION",
                "Factor_Usage": "HISTORICAL_SIMULATION",
                "Factor_Policy_ID": POLICY_ID,
                "Factor_Route": "HISTORICAL_SIMULATION_FACTOR",
                "Factor_Ready": "TRUE",
                "Boundary_Policy": BOUNDARY_ID,
                "Boundary_Ready": "TRUE",
                "Emission_kgCO2e": format(emission_kg, "f"),
                "Emission_tCO2e": format(emission_kg / Decimal("1000"), "f"),
                "Emission_Ready": "TRUE",
                "Calculation_QC": "PASS" if not blocking else "BLOCKED",
                "Governance_QC": "WARNING" if governance_warning else _text(item.get("Governance_QC")) or "PASS",
                "QC_Status": "WARNING" if governance_warning else "PASS",
                "Blocking_Codes": blocking,
                "Simulation_Flag": "TRUE",
                "Production_Eligible": "FALSE",
                "Overall_Status": "PASS_WITH_WARNING" if governance_warning else "PASS",
                "Ledger_Source_File": evidence.get("Source_File"),
                "Ledger_Source_SHA256": evidence.get("Source_SHA256"),
            }
        )
        applied += 1
        output.append(item)
    return output, {
        "status": "PASS" if applied else "NOT_APPLIED",
        "policy_id": POLICY_ID,
        "boundary_policy_id": BOUNDARY_ID,
        "factor_value": format(factor, "f"),
        "factor_unit": "kgCO2e/kg",
        "applied_records": applied,
        "skipped_non_cell_records": skipped_non_cell,
        "skipped_out_of_scope_records": skipped_scope,
        "skipped_activity_records": skipped_activity,
        "simulation_flag": True,
        "production_eligible": False,
    }


def reconcile_with_ledger(
    rows: list[dict[str, Any]], evidence: dict[str, Any] | None
) -> dict[str, Any]:
    activity = sum(
        (_decimal(row.get("Activity_Data_kg")) or Decimal(0) for row in rows),
        Decimal(0),
    )
    emission_t = sum(
        (_decimal(row.get("Emission_tCO2e")) or Decimal(0) for row in rows),
        Decimal(0),
    )
    expected_activity = _decimal((evidence or {}).get("Historical_Activity_kg"))
    expected_emission = _decimal((evidence or {}).get("Historical_Emission_tCO2e"))
    activity_difference = activity - expected_activity if expected_activity is not None else None
    emission_difference = emission_t - expected_emission if expected_emission is not None else None
    activity_pass = activity_difference is not None and abs(activity_difference) <= Decimal("0.000001")
    emission_pass = emission_difference is not None and abs(emission_difference) <= Decimal("0.000000001")
    return {
        "schema_version": "WP6_8_5_LEDGER_RECONCILIATION_V1",
        "status": "PASS" if activity_pass and emission_pass else "WARNING",
        "Record_Count": len(rows),
        "Calculated_Activity_kg": format(activity, "f"),
        "Historical_Activity_kg": "" if expected_activity is None else format(expected_activity, "f"),
        "Activity_Difference_kg": "" if activity_difference is None else format(activity_difference, "f"),
        "Activity_Reconciliation": "PASS" if activity_pass else "WARNING",
        "Calculated_Emission_tCO2e": format(emission_t, "f"),
        "Historical_Emission_tCO2e": "" if expected_emission is None else format(expected_emission, "f"),
        "Emission_Difference_tCO2e": "" if emission_difference is None else format(emission_difference, "f"),
        "Emission_Reconciliation": "PASS" if emission_pass else "WARNING",
        "Historical_Result_Used_As_Input": False,
    }
