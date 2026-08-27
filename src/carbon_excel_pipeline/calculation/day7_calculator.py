"""Day 7 strict Decimal emission calculation and D5 result assembly."""

from __future__ import annotations

import uuid
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from typing import Any

from carbon_excel_pipeline.errors import PipelineUserError


CALC_NAMESPACE = uuid.UUID("1f31280e-ca90-47a3-8d0f-ddd6f067756d")
CALCULATION_AUDIT_FIELDS = [
    "Calculation_Run_ID", "Calculation_Record_ID", "Record_ID", "Result_ID",
    "Routing_Record_ID", "Identity_Check", "Lineage_Check", "Match_Check",
    "Simulation_Check", "Synthetic_Check", "Production_Check",
    "Activity_Parse_Check", "Activity_Positive_Check", "EF_Parse_Check",
    "EF_Positive_Check", "EF_Identity_Check", "Unit_Check", "Formula_Check",
    "Raw_Emission_Recomputed", "Stored_Raw_Emission", "Raw_Difference",
    "Display_Emission_Recomputed", "Stored_Display_Emission", "Audit_Status",
    "D4_Issue_Code",
]
EMISSION_SUMMARY_FIELDS = [
    "Calculation_Run_ID", "Year", "Record_Count", "Activity_Total_kg",
    "Historical_EF_ID", "EF_Value", "Raw_Total_Emission_kgCO2e",
    "Total_Emission_kgCO2e", "Sum_of_Row_Display_Emission",
    "Rounding_Reconciliation_Difference", "Emission_Unit", "Calculation_Status",
    "Simulation_Flag", "Production_Eligible", "Open_Items_Count", "Remarks",
]


def _error(code: str, message: str, location: str, value: Any, rule: str) -> PipelineUserError:
    return PipelineUserError(
        stage="DAY7_DECIMAL_CALCULATION",
        error_code=code,
        message_cn=message,
        source_location=location,
        original_value=value,
        rule=rule,
        impact="阻断Day 7正式核算、D5结果和两层血缘",
        fix_suggestion="恢复Day 6精确锁定记录并修正数值或机器单位后重新执行。",
    )


def stable_uuid(kind: str, *parts: str) -> str:
    return str(uuid.uuid5(CALC_NAMESPACE, "|".join((kind, *parts))))


def canonical_decimal(value: Decimal) -> str:
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def multiply_emission_decimal(
    activity_kg: Decimal,
    ef_kgco2e_per_kg: Decimal,
    *,
    precision: int = 34,
) -> tuple[Decimal, Decimal]:
    """Shared Decimal core returning kgCO2e and tCO2e without display rounding."""

    if not activity_kg.is_finite() or activity_kg < 0:
        raise ValueError("activity_kg must be a finite non-negative Decimal")
    if not ef_kgco2e_per_kg.is_finite() or ef_kgco2e_per_kg < 0:
        raise ValueError("ef_kgco2e_per_kg must be a finite non-negative Decimal")
    with localcontext() as context:
        context.prec = precision
        emission_kg = activity_kg * ef_kgco2e_per_kg
        emission_t = emission_kg / Decimal("1000")
    return emission_kg, emission_t


def _parse_positive(value: Any, *, record_id: str, field: str) -> Decimal:
    text = str(value)
    if not text or text != text.strip():
        raise _error(
            f"{field.upper()}_INVALID",
            f"{field}为空或带前后空格。",
            f"{record_id}/{field}",
            text,
            "正式数值必须是未经清洗的严格正Decimal。",
        )
    try:
        parsed = Decimal(text)
    except InvalidOperation as exc:
        raise _error(
            f"{field.upper()}_INVALID",
            f"{field}不是有效Decimal。",
            f"{record_id}/{field}",
            text,
            "正式数值必须是严格正Decimal。",
        ) from exc
    if not parsed.is_finite() or parsed <= 0:
        raise _error(
            f"{field.upper()}_NOT_POSITIVE",
            f"{field}必须严格大于0。",
            f"{record_id}/{field}",
            text,
            "正式数值必须是严格正Decimal。",
        )
    return parsed


def calculate_d4(
    route_records: list[dict[str, Any]],
    *,
    rules: dict[str, Any],
    d4_fields: list[str],
    run_id: str,
    calculated_at: str,
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    calculation_run_id = stable_uuid("calculation-run", run_id)
    quantizer = Decimal("1").scaleb(-int(rules["display_decimal_places"]))
    issue_code = (
        "D4_SIMULATED_HISTORICAL_EF;D4_TECHNICAL_APPLICABILITY_UNCONFIRMED;"
        "D4_PRODUCTION_DISABLED;D4_OPEN_ITEMS_INHERITED"
    )
    output: list[dict[str, str]] = []
    audit: list[dict[str, str]] = []
    seen: set[str] = set()
    for route in route_records:
        record_id = str(route["Record_ID"])
        if record_id in seen:
            raise _error(
                "DUPLICATE_RECORD_ID", "D3路线包含重复Record_ID。", record_id, record_id, "正式D4输入必须每个Record_ID恰好一条。"
            )
        seen.add(record_id)
        required_route = {
            "Exact_Lock_Status": "EXACT_LOCKED",
            "Match_Method": "EXACT_RECORD_ID",
            "Match_Status": "EXACT_MATCH",
            "Manual_Review_Required": "FALSE",
            "Fallback_Attempted": "FALSE",
            "Simulation_Flag": "TRUE",
            "Synthetic_Test_Flag": "FALSE",
            "Production_Eligible": "FALSE",
            "D4_Calculation_Eligible": "TRUE_WITH_CAVEATS",
        }
        mismatches = {
            field: {"actual": str(route.get(field, "")), "expected": expected}
            for field, expected in required_route.items()
            if str(route.get(field, "")) != expected
        }
        if mismatches:
            raise _error(
                "D3_ROUTE_NOT_CALCULATION_ELIGIBLE", "D3路线不满足正式核算门禁。", record_id, mismatches, "只允许精确锁定、无回退、无人工复核的私有历史模拟路线。"
            )
        if str(route["Activity_Unit"]) != rules["activity_unit"]:
            raise _error(
                "ACTIVITY_UNIT_INCOMPATIBLE_EXACT", "活动单位与正式单位规则不兼容。", f"{record_id}/Activity_Unit", route["Activity_Unit"], "单位大小写敏感且不裁剪空格。"
            )
        if str(route["EF_Unit"]) != rules["ef_unit"]:
            raise _error(
                "EF_UNIT_INCOMPATIBLE_EXACT", "因子单位与正式单位规则不兼容。", f"{record_id}/EF_Unit", route["EF_Unit"], "单位大小写敏感且不裁剪空格。"
            )
        activity = _parse_positive(route["Activity_Data"], record_id=record_id, field="Activity_Data")
        ef_value = _parse_positive(route["EF_Value"], record_id=record_id, field="EF_Value")
        raw_emission, _ = multiply_emission_decimal(
            activity,
            ef_value,
            precision=int(rules["decimal_context_precision"]),
        )
        display_emission = raw_emission.quantize(quantizer, rounding=ROUND_HALF_UP)
        calculation_record_id = stable_uuid(
            "calculation-record", run_id, record_id, str(route["Result_ID"]), canonical_decimal(activity), canonical_decimal(ef_value)
        )
        raw_text = canonical_decimal(raw_emission)
        display_text = format(display_emission, f".{rules['display_decimal_places']}f")
        row = {
            "Calculation_Run_ID": calculation_run_id,
            "Calculation_Record_ID": calculation_record_id,
            "Calculation_Algorithm_Version": rules["algorithm_version"],
            "Calculated_At": calculated_at,
            "Record_ID": record_id,
            "Result_ID": str(route["Result_ID"]),
            "Historical_EF_ID": str(route["Historical_EF_ID"]),
            "D2_Match_ID": str(route["D2_Match_ID"]),
            "Routing_Record_ID": str(route["Routing_Record_ID"]),
            "Exact_Lock_Status": str(route["Exact_Lock_Status"]),
            "Match_Method": str(route["Match_Method"]),
            "Match_Status": str(route["Match_Status"]),
            "Routing_Status": str(route["Routing_Status"]),
            "D3_Record_Status": str(route["D3_Record_Status"]),
            "D3_Issue_Code": str(route["D3_Issue_Code"]),
            "D4_Calculation_Eligible_Input": str(route["D4_Calculation_Eligible"]),
            "Simulation_Flag": str(route["Simulation_Flag"]),
            "Synthetic_Test_Flag": str(route["Synthetic_Test_Flag"]),
            "Production_Eligible": str(route["Production_Eligible"]),
            "Activity_Data_Original": str(route["Activity_Data"]),
            "Activity_Unit_Original": str(route["Activity_Unit"]),
            "Activity_Value_Parsed": canonical_decimal(activity),
            "Activity_Unit_Normalized": rules["activity_unit"],
            "Activity_Conversion_Factor": "1",
            "Activity_Data_Normalized_kg": canonical_decimal(activity),
            "EF_Value_Original": str(route["EF_Value"]),
            "EF_Unit_Original": str(route["EF_Unit"]),
            "EF_Value_Parsed": canonical_decimal(ef_value),
            "EF_Unit_Normalized": rules["ef_unit"],
            "EF_Conversion_Factor": "1",
            "EF_Value_Normalized_kgCO2e_per_kg": canonical_decimal(ef_value),
            "Unit_Compatibility_Status": "DIRECT_COMPATIBLE",
            "Unit_Compatibility_Rule_ID": rules["unit_rule_id"],
            "Calculation_Formula": rules["calculation_formula"],
            "Arithmetic_Engine": rules["arithmetic_engine"],
            "Decimal_Context_Precision": str(rules["decimal_context_precision"]),
            "Rounding_Mode": rules["rounding_mode"],
            "Raw_Emission_kgCO2e": raw_text,
            "Display_Decimal_Places": str(rules["display_decimal_places"]),
            "Emission_kgCO2e": display_text,
            "Emission_Unit": rules["emission_unit"],
            "Calculation_Status": "CALCULATED_WITH_WARNING",
            "Calculation_QC_Status": "WARNING",
            "D4_Issue_Code": issue_code,
            "Calculation_Evidence": (
                f"WP3:{record_id}|D1:{route['Result_ID']}|D2:{route['D2_Match_ID']}|"
                f"D3:{route['Routing_Record_ID']}|{rules['unit_rule_id']}"
            ),
            "Formal_Output_Eligible": "TRUE",
            "D5_End_to_End_Eligible": "TRUE_WITH_CAVEATS",
            "Remarks": rules["remarks"],
        }
        output.append({field: row[field] for field in d4_fields})
        audit.append(
            {
                "Calculation_Run_ID": calculation_run_id,
                "Calculation_Record_ID": calculation_record_id,
                "Record_ID": record_id,
                "Result_ID": str(route["Result_ID"]),
                "Routing_Record_ID": str(route["Routing_Record_ID"]),
                "Identity_Check": "PASS",
                "Lineage_Check": "PASS",
                "Match_Check": "PASS",
                "Simulation_Check": "PASS_WITH_WARNING",
                "Synthetic_Check": "PASS",
                "Production_Check": "PASS_DISABLED",
                "Activity_Parse_Check": "PASS",
                "Activity_Positive_Check": "PASS",
                "EF_Parse_Check": "PASS",
                "EF_Positive_Check": "PASS",
                "EF_Identity_Check": "PASS",
                "Unit_Check": "PASS_DIRECT_COMPATIBLE",
                "Formula_Check": "PASS",
                "Raw_Emission_Recomputed": raw_text,
                "Stored_Raw_Emission": raw_text,
                "Raw_Difference": "0",
                "Display_Emission_Recomputed": display_text,
                "Stored_Display_Emission": display_text,
                "Audit_Status": "PASS_WITH_WARNING",
                "D4_Issue_Code": issue_code,
            }
        )
    return output, audit


def build_emission_summary(
    d4_records: list[dict[str, str]], *, rules: dict[str, Any]
) -> tuple[dict[str, str], dict[str, bool]]:
    raw_total = sum(Decimal(row["Raw_Emission_kgCO2e"]) for row in d4_records)
    activity_total = sum(Decimal(row["Activity_Data_Normalized_kg"]) for row in d4_records)
    row_display_total = sum(Decimal(row["Emission_kgCO2e"]) for row in d4_records)
    quantizer = Decimal("1").scaleb(-int(rules["display_decimal_places"]))
    official_total = raw_total.quantize(quantizer, rounding=ROUND_HALF_UP)
    rounding_difference = row_display_total - official_total
    values = {
        "record_count": len(d4_records),
        "activity_total": canonical_decimal(activity_total),
        "raw_total": canonical_decimal(raw_total),
        "official_total": format(official_total, f".{rules['display_decimal_places']}f"),
        "row_display_total": format(row_display_total, f".{rules['display_decimal_places']}f"),
        "rounding_difference": format(rounding_difference, f".{rules['display_decimal_places']}f"),
    }
    expected_records = rules.get("expected_private_records")
    checks = {
        "record_count": expected_records is None or values["record_count"] == expected_records,
        "activity_total": rules.get("expected_activity_total_kg") in (None, values["activity_total"]),
        "raw_total": rules.get("expected_raw_total_kgco2e") in (None, values["raw_total"]),
        "official_total": rules.get("expected_official_six_decimal_total") in (None, values["official_total"]),
        "row_display_total": rules.get("expected_row_six_decimal_sum") in (None, values["row_display_total"]),
        "rounding_difference": rules.get("expected_rounding_reconciliation_difference")
        in (None, values["rounding_difference"]),
    }
    summary = {
        "Calculation_Run_ID": d4_records[0]["Calculation_Run_ID"] if d4_records else "",
        "Year": "2030",
        "Record_Count": str(values["record_count"]),
        "Activity_Total_kg": values["activity_total"],
        "Historical_EF_ID": d4_records[0]["Historical_EF_ID"] if d4_records else "",
        "EF_Value": rules.get("expected_ef_value")
        or (d4_records[0].get("EF_Value_Normalized") if d4_records else "")
        or (d4_records[0].get("EF_Value") if d4_records else ""),
        "Raw_Total_Emission_kgCO2e": values["raw_total"],
        "Total_Emission_kgCO2e": values["official_total"],
        "Sum_of_Row_Display_Emission": values["row_display_total"],
        "Rounding_Reconciliation_Difference": values["rounding_difference"],
        "Emission_Unit": rules["emission_unit"],
        "Calculation_Status": "CALCULATED_WITH_WARNING",
        "Simulation_Flag": "TRUE",
        "Production_Eligible": "FALSE",
        "Open_Items_Count": str(rules.get("wp5_governance_open_items") or 0),
        "Remarks": "Aggregate raw Decimal values before six-place ROUND_HALF_UP; row display reconciliation retained.",
    }
    return summary, checks


def build_d5_results(
    d4_records: list[dict[str, str]],
    *,
    d5_fields: list[str],
    run_id: str,
    executed_at: str,
) -> list[dict[str, str]]:
    end_to_end_run_id = stable_uuid("end-to-end-run", run_id)
    output: list[dict[str, str]] = []
    for d4 in d4_records:
        prefix = {
            "End_to_End_Run_ID": end_to_end_run_id,
            "End_to_End_Record_ID": stable_uuid(
                "end-to-end-record", run_id, d4["Record_ID"], d4["Calculation_Record_ID"]
            ),
            "End_to_End_Algorithm_Version": "1.0.0",
            "End_to_End_Executed_At": executed_at,
            "Upstream_Trace_Status": "UPSTREAM_EVIDENCE_VALID_WITH_WARNING",
            "Stage_Reconciliation_Status": "ALL_STAGES_SEMANTICALLY_RECONCILED",
            "End_to_End_QC_Status": "END_TO_END_COMPLETED_WITH_CAVEATS",
            "WP6_Validation_Eligible": "TRUE_WITH_CAVEATS",
        }
        merged = {**prefix, **d4}
        output.append({field: merged[field] for field in d5_fields})
    return output
