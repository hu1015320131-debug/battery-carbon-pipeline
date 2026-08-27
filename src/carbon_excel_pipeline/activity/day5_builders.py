"""Build Decimal activity data, third-party input and layered WP3 open items."""

from __future__ import annotations

from decimal import Decimal
from typing import Any

from carbon_excel_pipeline.cleaning.raw_cleaner import canonical_decimal


RECORD_OPEN_ITEM_FIELDS = [
    "Open_Item_ID",
    "Record_ID",
    "Data_Quality_Flag",
    "Source_Issue_Code",
    "Missing_Fields",
    "Risk_Description",
    "Current_Decision",
    "Impact",
    "Recommended_Action",
    "Owner",
    "Target_Date",
    "Status",
    "Evidence",
]


def build_activity_records(
    checked_records: list[dict[str, Any]],
    *,
    activity_fields: list[str],
    semantic_zero_tolerance_g: Decimal = Decimal("0"),
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    output: list[dict[str, Any]] = []
    blocked: list[dict[str, Any]] = []
    for record in checked_records:
        if record["QC_Status"] == "ERROR":
            blocked.append(
                {
                    "Record_ID": record["Record_ID"],
                    "Source_Row": record["Source_Row"],
                    "QC_Status": record["QC_Status"],
                    "Issue_Code": record["Issue_Code"],
                    "Block_Reason": "UPSTREAM_QC_ERROR",
                }
            )
            continue
        pcs = Decimal(str(record["PCS"]))
        weight = Decimal(str(record["Unit_Weight_g"]))
        original = Decimal(str(record["Original_Activity_Value"]))
        total_g = pcs * weight
        total_kg = total_g / Decimal("1000")
        total_t = total_g / Decimal("1000000")
        diff_g = total_g - original
        if abs(diff_g) <= semantic_zero_tolerance_g:
            diff_g = Decimal("0")
        diff_rate = diff_g / original
        extended = {
            **record,
            "Total_Weight_g": canonical_decimal(total_g),
            "Total_Weight_kg": canonical_decimal(total_kg),
            "Total_Weight_t": canonical_decimal(total_t),
            "Activity_Diff_g": canonical_decimal(diff_g),
            "Activity_Diff_Rate": canonical_decimal(diff_rate),
        }
        output.append({field: extended[field] for field in activity_fields})
    return output, blocked


def _missing_third_party_fields(record: dict[str, Any]) -> list[str]:
    missing: list[str] = []
    if record.get("Project_Code") in (None, "", "UNKNOWN", "PENDING"):
        missing.append("Project_Code")
    if record.get("Cell_Model") in (None, "", "UNKNOWN", "PENDING"):
        missing.append("Cell_Model")
    if record.get("Chemistry") in (None, "", "UNKNOWN", "PENDING"):
        missing.append("Chemistry")
    return missing


def build_third_party_records(
    activity_records: list[dict[str, Any]],
    *,
    contract: dict[str, Any],
) -> list[dict[str, Any]]:
    constants = contract["constants"]
    output: list[dict[str, Any]] = []
    for record in activity_records:
        missing = _missing_third_party_fields(record)
        qc_status = record["QC_Status"]
        issue_code = record["Issue_Code"]
        remarks = ""
        if qc_status != "PASS":
            remarks = contract["warning_remarks_template"].format(
                qc=qc_status, issue=issue_code
            )
        result = {
            "Record_ID": record["Record_ID"],
            "Year": record["Year"],
            "Product_Type": constants["Product_Type"],
            "Product_Description": record["Product_Description"],
            "Activity_Data": record["Total_Weight_kg"],
            "Activity_Unit": constants["Activity_Unit"],
            "Supplier_Name": record["Supplier_Name"],
            "Project_Code": record["Project_Code"],
            "Cell_Model": record["Cell_Model"],
            "Chemistry": record["Chemistry"],
            "PCS": record["PCS"],
            "Unit_Weight_g": record["Unit_Weight_g"],
            "Production_Country": constants["Production_Country"],
            "Factory": constants["Factory"],
            "Technology": constants["Technology"],
            "Supplier_PCF_Available": constants["Supplier_PCF_Available"],
            "Evidence": constants["Evidence"],
            "Data_Quality_Flag": qc_status,
            "Missing_Fields": ";".join(missing) if missing else "NONE",
            "Remarks": remarks,
        }
        output.append({field: result[field] for field in contract["fields"]})
    return output


def build_record_open_items(
    activity_records: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    warning_records = [
        record for record in activity_records if record["QC_Status"] == "WARNING"
    ]
    for index, record in enumerate(warning_records, start=1):
        missing = _missing_third_party_fields(record)
        output.append(
            {
                "Open_Item_ID": f"WP3-REC-{index:04d}",
                "Record_ID": record["Record_ID"],
                "Data_Quality_Flag": record["QC_Status"],
                "Source_Issue_Code": record["Issue_Code"],
                "Missing_Fields": ";".join(missing) if missing else "NONE",
                "Risk_Description": "Record-level master data is incomplete.",
                "Current_Decision": "ALLOW_WITH_WARNING",
                "Impact": "Retain warning in external assessment input.",
                "Recommended_Action": "Complete governed master data before final assessment.",
                "Owner": "PENDING",
                "Target_Date": "",
                "Status": "OPEN",
                "Evidence": "DAY5_AUTOMATED_QC",
            }
        )
    return output
