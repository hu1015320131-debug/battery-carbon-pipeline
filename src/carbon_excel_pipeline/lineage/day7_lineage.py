"""Build immutable-schema WP5 lineage and separate Demo extended lineage."""

from __future__ import annotations

from typing import Any

from carbon_excel_pipeline.errors import PipelineUserError


def _error(code: str, value: Any, message: str = "两层血缘Record_ID无法一一对应。") -> PipelineUserError:
    return PipelineUserError(
        stage="DAY7_TWO_LAYER_LINEAGE",
        error_code=code,
        message_cn=message,
        source_location="Record_ID",
        original_value=value,
        rule="D1/D2/D3/D4/D5、活动数据、冻结血缘和扩展血缘必须按Record_ID一一对应且无孤儿。",
        impact="阻断Day 7两层血缘和端到端结果",
        fix_suggestion="Restore unique Record_IDs in every stage and rerun.",
    )


def _unique_map(records: list[dict[str, Any]], label: str) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for row in records:
        record_id = str(row.get("Record_ID", ""))
        if not record_id or record_id in result:
            raise _error("LINEAGE_ID_NOT_UNIQUE", {"stage": label, "record_id": record_id})
        result[record_id] = row
    return result


def build_two_layer_lineage(
    *,
    activity_records: list[dict[str, Any]],
    d1_records: list[dict[str, Any]],
    d2_records: list[dict[str, Any]],
    d3_records: list[dict[str, Any]],
    d4_records: list[dict[str, Any]],
    d5_records: list[dict[str, Any]],
    frozen_fields: list[str],
    extended_fields: list[str],
    metadata: dict[str, str],
) -> tuple[list[dict[str, str]], list[dict[str, str]], dict[str, Any]]:
    maps = {
        "activity": _unique_map(activity_records, "activity"),
        "d1": _unique_map(d1_records, "d1"),
        "d2": _unique_map(d2_records, "d2"),
        "d3": _unique_map(d3_records, "d3"),
        "d4": _unique_map(d4_records, "d4"),
        "d5": _unique_map(d5_records, "d5"),
    }
    id_sets = {label: set(records) for label, records in maps.items()}
    reference = id_sets["activity"]
    if any(values != reference for values in id_sets.values()):
        raise _error(
            "LINEAGE_ID_SET_MISMATCH",
            {label: {"count": len(values), "missing": len(reference - values), "orphan": len(values - reference)} for label, values in id_sets.items()},
        )

    frozen: list[dict[str, str]] = []
    extended: list[dict[str, str]] = []
    for activity in activity_records:
        record_id = str(activity["Record_ID"])
        d1 = maps["d1"][record_id]
        d2 = maps["d2"][record_id]
        d3 = maps["d3"][record_id]
        d4 = maps["d4"][record_id]
        d5 = maps["d5"][record_id]
        frozen_row = {
            "End_to_End_Run_ID": str(d5["End_to_End_Run_ID"]),
            "End_to_End_Record_ID": str(d5["End_to_End_Record_ID"]),
            "Record_ID": record_id,
            "Activity_Source_File": metadata["activity_source_file"],
            "Activity_Source_SHA256": metadata["activity_source_sha256"],
            "Third_Party_Input_Source_File": metadata["third_party_source_file"],
            "Third_Party_Input_Source_SHA256": metadata["third_party_source_sha256"],
            "Historical_EF_ID": str(d1["Historical_EF_ID"]),
            "Historical_Relationship_ID": str(d1["Historical_Relationship_ID"]),
            "D1_Result_ID": str(d1["Result_ID"]),
            "D1_Record_Status": str(d1["D1_Record_Status"]),
            "D1_Simulation_Flag": str(d1["Simulation_Flag"]),
            "D1_Production_Eligible": str(d1["Production_Eligible"]),
            "D2_Match_Run_ID": str(d2["Match_Run_ID"]),
            "D2_Match_ID": str(d2["Match_ID"]),
            "D2_Match_Method": str(d2["Match_Method"]),
            "D2_Match_Status": str(d2["Match_Status"]),
            "D2_Lineage_Check_Status": str(d2["Lineage_Check_Status"]),
            "D3_Routing_Run_ID": str(d3["Routing_Run_ID"]),
            "D3_Routing_Record_ID": str(d3["Routing_Record_ID"]),
            "D3_Exact_Lock_Status": str(d3["Exact_Lock_Status"]),
            "D3_Routing_Status": str(d3["Routing_Status"]),
            "D3_Manual_Review_Required": str(d3["Manual_Review_Required"]),
            "D4_Calculation_Run_ID": str(d4["Calculation_Run_ID"]),
            "D4_Calculation_Record_ID": str(d4["Calculation_Record_ID"]),
            "D4_Calculation_Status": str(d4["Calculation_Status"]),
            "Activity_Data_Normalized_kg": str(d4["Activity_Data_Normalized_kg"]),
            "EF_Value_Normalized_kgCO2e_per_kg": str(d4["EF_Value_Normalized_kgCO2e_per_kg"]),
            "Raw_Emission_kgCO2e": str(d4["Raw_Emission_kgCO2e"]),
            "Emission_kgCO2e": str(d4["Emission_kgCO2e"]),
            "Emission_Unit": str(d4["Emission_Unit"]),
            "End_to_End_Trace_Status": "TRACE_COMPLETE_WITH_CAVEATS",
        }
        frozen.append({field: frozen_row[field] for field in frozen_fields})
        extended_row = {
            "Record_ID": record_id,
            "End_to_End_Run_ID": str(d5["End_to_End_Run_ID"]),
            "End_to_End_Record_ID": str(d5["End_to_End_Record_ID"]),
            "Raw_Input_Source_File": metadata["raw_input_source_file"],
            "Raw_Input_SHA256": metadata["raw_input_sha256"],
            "Received_Input_Copy": metadata["received_input_copy"],
            "Received_Input_SHA256": metadata["received_input_sha256"],
            "Import_Run_ID": metadata["import_run_id"],
            "Source_Sheet": str(activity["Source_Sheet"]),
            "Source_Row": str(activity["Source_Row"]),
            "Day3_Candidate_File": "02_scope_filter/candidate_records.csv",
            "Day3_Exclusion_Audit_File": "02_scope_filter/excluded_records_audit.csv",
            "Day3_Disposition": "IN_SCOPE_CANDIDATE",
            "Day4_Standard_File": "03_standardized/day4_standard_31_fields.csv",
            "Day5_Activity_File": "05_activity/day5_activity_36_fields.csv",
            "Day6_D1_File": "07_factor_results/day6_d1_factor_results_45_fields.csv",
            "Day6_D2_File": "08_matching/day6_d2_exact_matches_57_fields.csv",
            "Day6_D3_File": "08_matching/day6_d3_exact_routes_36_fields.csv",
            "Day7_D4_File": "09_calculation/day7_d4_calculation_48_fields.csv",
            "Day7_D5_File": "10_output/day7_d5_end_to_end_56_fields.csv",
            "Profile_ID": metadata["profile_id"],
            "Profile_Config_SHA256": metadata["profile_config_sha256"],
            "Calculation_Config_SHA256": metadata["calculation_config_sha256"],
            "Frozen_Lineage_Join_Status": "ONE_TO_ONE",
            "Extended_Trace_Status": "TRACE_COMPLETE",
        }
        extended.append({field: extended_row[field] for field in extended_fields})

    frozen_ids = [row["Record_ID"] for row in frozen]
    extended_ids = [row["Record_ID"] for row in extended]
    qa = {
        "status": "PASS" if frozen_ids == extended_ids and len(frozen_ids) == len(set(frozen_ids)) else "FAIL",
        "frozen_field_count": len(frozen_fields),
        "extended_field_count": len(extended_fields),
        "frozen_records": len(frozen),
        "extended_records": len(extended),
        "record_id_order_equal": frozen_ids == extended_ids,
        "record_ids_unique": len(frozen_ids) == len(set(frozen_ids)) == len(extended_ids) == len(set(extended_ids)),
        "missing_extended_records": len(set(frozen_ids) - set(extended_ids)),
        "orphan_extended_records": len(set(extended_ids) - set(frozen_ids)),
        "frozen_schema_mutated": False,
        "extended_lineage_described_as_wp5_frozen": False,
    }
    return frozen, extended, qa
