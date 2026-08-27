"""Day 5 upstream quality checks over the 31-field standard records."""

from __future__ import annotations

from collections import Counter
from decimal import Decimal, InvalidOperation
from typing import Any


def _positive_decimal(value: Any) -> Decimal | None:
    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None
    return result if result.is_finite() and result > 0 else None


def _positive_integer(value: Any) -> Decimal | None:
    result = _positive_decimal(value)
    if result is None or result != result.to_integral_value():
        return None
    return result


def run_quality_checks(
    records: list[dict[str, Any]],
    *,
    config: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    id_counts = Counter(str(record.get("Record_ID", "")) for record in records)
    checked: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    warning_order = config["warning_issue_order"]
    reconciliation_tolerance = Decimal(
        config.get("registered_float_tail_tolerance_g", "0")
    )

    for record in records:
        error_codes: list[str] = []
        warning_codes: list[str] = []
        record_id = str(record.get("Record_ID", ""))

        missing_fields = [
            field
            for field in config["required_fields"]
            if record.get(field) is None or record.get(field) == ""
        ]
        if missing_fields:
            error_codes.append("REQUIRED_FIELD_MISSING")
        if not record_id:
            error_codes.append("RECORD_ID_MISSING")
        elif id_counts[record_id] != 1:
            error_codes.append("RECORD_ID_DUPLICATE")

        try:
            if int(record["Year"]) <= 0:
                raise ValueError
        except (KeyError, TypeError, ValueError):
            error_codes.append("YEAR_INVALID")
        try:
            if int(record["Source_Row"]) <= 0:
                raise ValueError
        except (KeyError, TypeError, ValueError):
            error_codes.append("SOURCE_ROW_INVALID")

        pcs = _positive_integer(record.get("PCS"))
        weight = _positive_decimal(record.get("Unit_Weight_g"))
        original = _positive_decimal(record.get("Original_Activity_Value"))
        if pcs is None:
            error_codes.append("PCS_INVALID")
        if weight is None:
            error_codes.append("UNIT_WEIGHT_INVALID")
        if original is None:
            error_codes.append("ORIGINAL_ACTIVITY_INVALID")
        if record.get("Original_Activity_Unit") != config["formal_activity_unit"]:
            error_codes.append("ACTIVITY_UNIT_INVALID_EXACT")
        if pcs is not None and weight is not None and original is not None:
            if abs(pcs * weight - original) > reconciliation_tolerance:
                error_codes.append("ACTIVITY_RECONCILIATION_MISMATCH")

        if record.get("Supplier_Status") == "PENDING" or record.get(
            "Mapping_Status"
        ) != "SUPPLIER_MAPPED":
            error_codes.append("SUPPLIER_UNMAPPED")
        if record.get("Mapping_Source") in (None, "", "NO_MAPPING"):
            error_codes.append("MAPPING_SOURCE_MISSING")
        if record.get("Customer_Mapping_Status") != "CUSTOMER_MAPPED":
            warning_codes.append("CUSTOMER_UNMAPPED")
        if record.get("Project_Code") in (None, "", "PENDING", "UNKNOWN"):
            warning_codes.append("PROJECT_MAPPING_PENDING")
        if record.get("Cell_Model") in (None, "", "PENDING", "UNKNOWN"):
            warning_codes.append("CELL_MODEL_UNKNOWN")
        if record.get("Chemistry") in (None, "", "PENDING", "UNKNOWN"):
            warning_codes.append("CHEMISTRY_UNKNOWN")

        error_codes = list(dict.fromkeys(error_codes))
        warning_codes = [
            code for code in warning_order if code in set(warning_codes)
        ]
        issue_codes = error_codes + warning_codes
        status = "ERROR" if error_codes else "WARNING" if warning_codes else "PASS"
        updated = dict(record)
        updated["Data_Status"] = "COMPLETE" if status == "PASS" else "PARTIAL"
        updated["QC_Status"] = status
        updated["Issue_Code"] = "NONE" if not issue_codes else ";".join(issue_codes)
        checked.append(updated)

        for code in issue_codes:
            severity = "ERROR" if code in error_codes else "WARNING"
            issues.append(
                {
                    "Record_ID": record_id,
                    "Source_File": record.get("Source_File", ""),
                    "Source_Sheet": record.get("Source_Sheet", ""),
                    "Source_Row": record.get("Source_Row", ""),
                    "Rule_ID": f"DAY5-{code}",
                    "Severity": severity,
                    "Issue_Code": code,
                    "Missing_Required_Fields": ";".join(missing_fields)
                    if code == "REQUIRED_FIELD_MISSING"
                    else "",
                }
            )

    counts = Counter(record["QC_Status"] for record in checked)
    counts_complete = {name: counts.get(name, 0) for name in ("PASS", "WARNING", "ERROR")}
    summary = {
        "status_counts": counts_complete,
        "record_count": len(checked),
        "issue_record_count": sum(record["QC_Status"] != "PASS" for record in checked),
        "issue_detail_count": len(issues),
        "error_records_blocked": counts_complete["ERROR"],
    }
    return checked, issues, summary
