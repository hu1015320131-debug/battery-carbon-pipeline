from carbon_excel_pipeline.wp6_8.pipeline import run_live_delivery


def test_live_delivery_does_not_invent_independent_pass(tmp_path):
    run = tmp_path / "RUN-LIVE"
    run.mkdir()
    row = {
        "Record_ID": "2024-DY2-SYNA-DX000001",
        "Year": "2024",
        "Source_File": "a.xlsx",
        "Source_SHA256": "A" * 64,
        "Source_Sheet": "S",
        "Source_Row": "3",
        "Business_Unit": "二部",
        "Purchase_Category": "电芯",
        "Product_Description": "SYNA",
        "Activity_Data_kg": "1000",
        "Activity_Unit": "kg/year",
        "Activity_Method": "DIRECT_REPORTED_MASS",
        "EF_Value": "1",
        "EF_Unit": "kgCO2e/kg",
        "Emission_kgCO2e": "1000",
        "Emission_tCO2e": "1",
        "Run_ID": "RUN-LIVE",
        "Boundary_Ready": "TRUE",
        "Calculation_QC": "PASS",
        "Governance_QC": "WARNING",
        "Overall_Validation_Status": "NOT_RUN",
        "Simulation_Flag": "TRUE",
        "Production_Eligible": "FALSE",
    }
    result = run_live_delivery(
        run_dir=run,
        canonical_rows=[row],
        route_decision={"Activity_Route": "DIRECT_REPORTED_MASS", "Factor_Route": "SOURCE_EMBEDDED_FACTOR"},
        validation_rows=[],
        input_file="a.xlsx",
        input_sha256="A" * 64,
        recognition_summary={"recognition_status": "RECOGNIZED"},
        capability_summary={"activity_ready_count": 1},
        independent_status="NOT_RUN",
    )
    summary = (run / "08_download" / "run_summary.json").read_text(encoding="utf-8")
    live = (run / "08_download" / "wp6_8_live_summary.json").read_text(encoding="utf-8")
    assert "INDEPENDENT_CALCULATION_PASS" not in summary
    assert '"Independent_Validation_Status": "NOT_RUN"' in summary or '"Independent_Validation_Status": "NOT_RUN"' in live
    assert "NOT_AVAILABLE_FOR_SINGLE_INPUT_RUN" in live
    assert result["status"] in {"PASS", "PASS_WITH_WARNING", "PARTIAL_RESULT"}
