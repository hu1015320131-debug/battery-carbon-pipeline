"""Day 6 exact Record_ID matching and exact-lock routing."""

from __future__ import annotations

import uuid
from collections import Counter
from datetime import datetime, timezone
from typing import Any


MATCH_NAMESPACE = uuid.UUID("76195a2d-8767-4b9c-83e2-b9bb5a943428")
ANOMALY_FIELDS = [
    "Record_ID",
    "Activity_Candidate_Count",
    "Result_Candidate_Count",
    "Issue_Code",
    "Match_Status",
    "Fallback_Attempted",
]


def _stable_uuid(kind: str, *parts: str) -> str:
    return str(uuid.uuid5(MATCH_NAMESPACE, "|".join((kind, *parts))))


def exact_match_and_route(
    activity_records: list[dict[str, Any]],
    factor_records: list[dict[str, Any]],
    *,
    d2_fields: list[str],
    route_fields: list[str],
    run_id: str,
    evaluated_at: str | None = None,
) -> tuple[list[dict[str, str]], list[dict[str, str]], list[dict[str, Any]]]:
    """Return only unique 1:1 exact matches; every other key is an anomaly."""
    evaluated_at = evaluated_at or datetime.now(timezone.utc).isoformat()
    activity_counts = Counter(str(row["Record_ID"]) for row in activity_records)
    factor_counts = Counter(str(row["Record_ID"]) for row in factor_records)
    activities = {str(row["Record_ID"]): row for row in activity_records}
    factors = {str(row["Record_ID"]): row for row in factor_records}
    anomalies: list[dict[str, Any]] = []
    d2_output: list[dict[str, str]] = []
    route_output: list[dict[str, str]] = []
    match_run_id = _stable_uuid("match-run", run_id)
    routing_run_id = _stable_uuid("routing-run", run_id)

    for record_id in sorted(set(activity_counts) | set(factor_counts)):
        activity_count = activity_counts[record_id]
        result_count = factor_counts[record_id]
        if activity_count != 1 or result_count != 1:
            if activity_count == 1 and result_count == 0:
                issue = "UNMATCHED_FACTOR_RESULT"
            elif activity_count == 0 and result_count >= 1:
                issue = "ORPHAN_FACTOR_RESULT"
            elif activity_count > 1:
                issue = "DUPLICATE_ACTIVITY_RECORD_ID"
            else:
                issue = "MULTIPLE_FACTOR_CANDIDATES"
            anomalies.append(
                {
                    "Record_ID": record_id,
                    "Activity_Candidate_Count": activity_count,
                    "Result_Candidate_Count": result_count,
                    "Issue_Code": issue,
                    "Match_Status": "MANUAL_REVIEW_REQUIRED",
                    "Fallback_Attempted": "FALSE",
                }
            )
            continue

        activity = activities[record_id]
        factor = factors[record_id]
        match_id = _stable_uuid("match", run_id, record_id, str(factor["Result_ID"]))
        inherited_warning = str(activity["QC_Status"]) == "WARNING"
        d2_issue = (
            "D2_SIMULATION_INPUT_ONLY;D2_TECHNICAL_APPLICABILITY_UNCONFIRMED;"
            "D2_CONSTRUCTED_KEY_ALIGNMENT;D2_RESULT_WARNING_INHERITED"
        )
        if inherited_warning:
            d2_issue += ";D2_UPSTREAM_WARNING_INHERITED"
        d2_row = {
            "Match_Run_ID": match_run_id,
            "Match_ID": match_id,
            "Match_Algorithm_Version": "1.0.0",
            "Match_Method": "EXACT_RECORD_ID",
            "Match_Key_Field": "Record_ID",
            "Match_Key_Value": record_id,
            "Activity_Candidate_Count": "1",
            "Result_Candidate_Count": "1",
            "Match_Status": "EXACT_MATCH",
            "Match_QC_Status": "PASS",
            "D2_Record_Status": "WARNING",
            "D2_Issue_Code": d2_issue,
            "Matched_At": evaluated_at,
            "Record_ID": record_id,
            "Activity_Year": str(activity["Year"]),
            "Supplier_Abbreviation": str(activity["Supplier_Abbreviation"]),
            "Supplier_Name": str(activity["Supplier_Name"]),
            "Supplier_ID": str(activity["Supplier_ID"]),
            "Business_Unit": str(activity["Business_Unit"]),
            "Activity_Category": str(activity["Activity_Category"]),
            "Activity_Category_Code": str(activity["Activity_Category_Code"]),
            "Project_Code": str(activity["Project_Code"]),
            "Cell_Model": str(activity["Cell_Model"]),
            "Material_ID": str(activity["Material_ID"]),
            "Chemistry": str(activity["Chemistry"]),
            "Product_Description": str(activity["Product_Description"]),
            "Activity_Data": str(activity["Total_Weight_kg"]),
            "Activity_Unit": "kg/year",
            "Upstream_QC_Status": str(activity["QC_Status"]),
            "Upstream_Issue_Code": str(activity["Issue_Code"]),
            "Result_ID": str(factor["Result_ID"]),
            "Historical_EF_ID": str(factor["Historical_EF_ID"]),
            "Result_Type": str(factor["Result_Type"]),
            "EF_Value": str(factor["EF_Value"]),
            "EF_Unit": str(factor["EF_Unit"]),
            "EF_Source_Type": str(factor["EF_Source_Type"]),
            "Source_Name": str(factor["Source_Name"]),
            "Source_Version": str(factor["Source_Version"]),
            "Receipt_Source_Type": str(factor["Receipt_Source_Type"]),
            "Receipt_Source_File": str(factor["Receipt_Source_File"]),
            "Receipt_Source_SHA256": str(factor["Receipt_Source_SHA256"]),
            "Simulation_Flag": str(factor["Simulation_Flag"]),
            "Simulation_Type": str(factor["Simulation_Type"]),
            "Simulation_Source": str(factor["Simulation_Source"]),
            "Simulation_Purpose": str(factor["Simulation_Purpose"]),
            "Data_Truth_Class": str(factor["Data_Truth_Class"]),
            "Synthetic_Test_Flag": str(factor["Synthetic_Test_Flag"]),
            "Production_Eligible": str(factor["Production_Eligible"]),
            "D1_Record_Status": str(factor["D1_Record_Status"]),
            "D1_Issue_Code": str(factor["D1_Issue_Code"]),
            "Historical_Relationship_ID": str(factor["Historical_Relationship_ID"]),
            "Historical_Link_Method": str(factor["Historical_Link_Method"]),
            "Lineage_Check_Status": "PASS",
            "Exact_Match_Eligible": "TRUE",
            "D3_Fallback_Required": "FALSE",
            "D3_Routing_Status": "NO_FALLBACK_REQUIRED",
            "D2_Conclusion": "EXACT_MATCH_CONFIRMED_WITH_CAVEATS",
        }
        d2_output.append({field: d2_row[field] for field in d2_fields})

        route_row = {
            "Routing_Run_ID": routing_run_id,
            "Routing_Record_ID": _stable_uuid("route", run_id, record_id, str(factor["Result_ID"])),
            "Routing_Algorithm_Version": "1.0.0",
            "Evaluated_At": evaluated_at,
            "Record_ID": record_id,
            "Result_ID": str(factor["Result_ID"]),
            "D2_Match_ID": match_id,
            "D2_Match_Status": "EXACT_MATCH",
            "Exact_Lock_Status": "EXACT_LOCKED",
            "Fallback_Eligible": "FALSE",
            "Fallback_Attempted": "FALSE",
            "Fallback_Tier_Selected": "",
            "Fallback_Key_Field": "",
            "Fallback_Key_Value": "",
            "Activity_Candidate_Count": "1",
            "Result_Candidate_Count": "1",
            "Cross_Tier_Conflict": "FALSE",
            "Fallback_Confidence_Score": "1",
            "Fallback_Confidence_Band": "EXACT",
            "Match_Method": "EXACT_RECORD_ID",
            "Match_Status": "EXACT_MATCH",
            "Routing_Status": "D4_EXACT_BASELINE_WITH_CAVEATS",
            "Manual_Review_Required": "FALSE",
            "D3_Record_Status": "WARNING",
            "D3_Issue_Code": (
                "D3_EXACT_MATCH_LOCKED;D3_FORMAL_FALLBACK_NOT_REQUIRED;"
                "D3_RESULT_FALLBACK_METADATA_UNAVAILABLE;"
                "D3_TECHNICAL_APPLICABILITY_UNCONFIRMED;D3_NO_EMISSION_CALCULATION"
            ),
            "Simulation_Flag": str(factor["Simulation_Flag"]),
            "Synthetic_Test_Flag": str(factor["Synthetic_Test_Flag"]),
            "Production_Eligible": str(factor["Production_Eligible"]),
            "Activity_Data": str(activity["Total_Weight_kg"]),
            "Activity_Unit": "kg/year",
            "Historical_EF_ID": str(factor["Historical_EF_ID"]),
            "EF_Value": str(factor["EF_Value"]),
            "EF_Unit": str(factor["EF_Unit"]),
            "D4_Calculation_Eligible": "TRUE_WITH_CAVEATS",
            "D4_Gate_Route": "EXACT_MATCH_BASELINE",
            "D3_Conclusion": "EXACT_MATCH_PRESERVED_NO_FALLBACK_ATTEMPTED",
        }
        route_output.append({field: route_row[field] for field in route_fields})

    return d2_output, route_output, anomalies
