from __future__ import annotations

from decimal import Decimal

from carbon_excel_pipeline.wp6_8 import AUDIT_FIELDS, LEDGER_FIELDS, build_delivery_rows


def _canonical(record_id: str = "2025-DY2-SYNA-DX000001") -> dict[str, str]:
    return {
        "Record_ID": record_id,
        "Year": "2025",
        "Source_File": "synthetic.xlsx",
        "Source_Sheet": "Data",
        "Source_Row": "2",
        "Business_Unit": "二部",
        "Purchase_Category": "电芯",
        "Product_Description": "Synthetic Cell",
        "Activity_Data_kg": "1234.5678",
        "Activity_Method": "PCS_WEIGHT_DERIVED",
        "EF_Value": "1.250000",
        "EF_Unit": "kgCO2e/kg",
        "Emission_kgCO2e": "17099.3163892908",
        "Emission_tCO2e": "17.0993163892908",
        "Activity_Ready": "TRUE",
        "Emission_Ready": "TRUE",
        "Boundary_Ready": "TRUE",
        "Calculation_QC": "PASS",
        "Governance_QC": "WARNING",
        "Governance_Issue_Codes": "CHEMISTRY_UNKNOWN",
        "Simulation_Flag": "TRUE",
        "Production_Eligible": "FALSE",
    }


def test_suggested_ledger_and_audit_preserve_contract_and_reconcile() -> None:
    canonical = _canonical()
    result = build_delivery_rows(
        [canonical],
        enrichment_by_id={
            canonical["Record_ID"]: {
                "Activity_Category": "外购原料",
                "Original_Activity_Value": "1234567.8",
                "Original_Activity_Unit": "g/year",
            }
        },
        validation_by_id={
            canonical["Record_ID"]: {
                "Source_SHA256": "A" * 64,
                "Independent_EF_Source": "controlled synthetic factor",
            }
        },
        fallback_run_id="WP6-4-SYNTHETIC",
    )

    assert result["status"] == "PASS_WITH_WARNING"
    assert list(result["ledger_rows"][0]) == LEDGER_FIELDS
    assert list(result["audit_rows"][0]) == AUDIT_FIELDS
    ledger = result["ledger_rows"][0]
    audit = result["audit_rows"][0]
    assert Decimal(ledger["年度购买原料量（t/year）"]) * 1000 == Decimal(
        audit["Activity_Data_kg"]
    )
    assert Decimal(ledger["LCA排放因子（kgCO2e/kg）"]) == Decimal(audit["EF_Value"])
    assert Decimal(ledger["GHG排放量（tCO2e/year）"]) * 1000 == Decimal(
        audit["Emission_kgCO2e"]
    )
    assert audit["Source_SHA256"] == "A" * 64
    assert audit["Warning_Codes"] == "CHEMISTRY_UNKNOWN"


def test_partial_result_keeps_calculable_rows_and_explains_exclusion() -> None:
    blocked = _canonical("2025-DY2-SYNA-DX000002")
    blocked["Emission_kgCO2e"] = ""
    result = build_delivery_rows([_canonical(), blocked])
    assert result["status"] == "PARTIAL_RESULT"
    assert result["included_record_count"] == 1
    assert result["excluded_record_count"] == 1
    assert result["excluded_rows"][0]["Record_ID"] == "2025-DY2-SYNA-DX000002"
    assert "Emission_kgCO2e_MISSING" in result["excluded_rows"][0]["Reason_Codes"]


def test_fully_unusable_input_is_blocked_without_crashing() -> None:
    blocked = _canonical()
    blocked["Activity_Data_kg"] = ""
    result = build_delivery_rows([blocked])
    assert result["status"] == "BLOCKED"
    assert result["ledger_rows"] == []
    assert result["excluded_record_count"] == 1
