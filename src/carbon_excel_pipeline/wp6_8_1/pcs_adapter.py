"""Adapt a same-Run Day3–Day7 PCS result into live canonical rows."""

from __future__ import annotations

import csv
import json
from decimal import Decimal
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.wp6_4.quality import assess_quality_layers


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _index(rows: list[dict[str, str]]) -> dict[str, dict[str, str]]:
    return {row["Record_ID"]: row for row in rows}


def adapt_pcs_current_run(
    run_dir: Path,
    *,
    route_decision: dict[str, Any],
) -> list[dict[str, Any]]:
    """Read this Run's Day4/5/7 artifacts. Never open another Run directory."""

    run = run_dir.expanduser().resolve()
    receipt = _load_json(run / "01_import" / "file_receipt_report.json")
    standard = _index(_read_csv(run / "03_standardized" / "day4_standard_31_fields.csv"))
    activity = _index(_read_csv(run / "05_activity" / "day5_activity_36_fields.csv"))
    result = _read_csv(run / "10_output" / "day7_d5_end_to_end_56_fields.csv")
    input_sha = str(receipt.get("source_sha256", "")).upper()
    rows: list[dict[str, Any]] = []
    for current in result:
        record_id = current["Record_ID"]
        standard_row = standard[record_id]
        activity_row = activity[record_id]
        quality = assess_quality_layers(
            activity=activity_row, result=current, standard=standard_row
        )
        raw_emission = current.get("Raw_Emission_kgCO2e", "")
        activity_kg = current.get("Activity_Data_Normalized_kg", "")
        emission_t = (
            format(Decimal(raw_emission) / Decimal("1000"), "f") if raw_emission else ""
        )
        rows.append(
            {
                "Record_ID": record_id,
                "Year": standard_row.get("Year", ""),
                "Source_File": activity_row.get("Source_File", ""),
                "Source_SHA256": input_sha,
                "Source_Sheet": activity_row.get("Source_Sheet", ""),
                "Source_Row": activity_row.get("Source_Row", ""),
                "Business_Unit": standard_row.get("Business_Unit", ""),
                "Purchase_Type": standard_row.get("Activity_Category", ""),
                "Purchase_Category": standard_row.get("Activity_Category", ""),
                "Product_Description": standard_row.get("Product_Description", ""),
                "Original_Activity_Value": activity_row.get("Original_Activity_Value", ""),
                "Original_Activity_Unit": activity_row.get("Original_Activity_Unit", ""),
                "Quantity_PCS": activity_row.get("PCS", ""),
                "Unit_Weight": activity_row.get("Unit_Weight_g", ""),
                "Unit_Weight_Unit": "g/PCS",
                "Activity_Data_kg": activity_kg,
                "Activity_Unit": current.get("Activity_Unit_Normalized", "kg/year"),
                "Activity_Method": "PCS_WEIGHT_DERIVED",
                "EF_Value": current.get("EF_Value_Normalized_kgCO2e_per_kg", ""),
                "EF_Unit": current.get("EF_Unit_Normalized", "kgCO2e/kg"),
                "EF_Source": route_decision.get("Factor_Source") or current.get("EF_Source", ""),
                "EF_Usage": route_decision.get("Factor_Usage") or "HISTORICAL_SIMULATION",
                "Emission_kgCO2e": raw_emission,
                "Emission_tCO2e": emission_t,
                "Display_Emission_kgCO2e": current.get("Emission_kgCO2e", ""),
                "Chemistry": standard_row.get("Chemistry", ""),
                "Supplier": standard_row.get("Supplier_Name", ""),
                "Supplier_Name": standard_row.get("Supplier_Name", ""),
                "Project": standard_row.get("Project_Code", ""),
                "Project_Code": standard_row.get("Project_Code", ""),
                "Model": standard_row.get("Cell_Model", ""),
                "Cell_Model": standard_row.get("Cell_Model", ""),
                "QC_Status": activity_row.get("QC_Status", ""),
                "Warning_Codes": activity_row.get("Issue_Code", ""),
                "Run_ID": run.name,
                "Simulation_Flag": "TRUE",
                "Production_Eligible": "FALSE",
                "Factor_Policy_ID": route_decision.get("Factor_Policy_ID"),
                "Factor_Route": route_decision.get("Factor_Route"),
                "Boundary_Policy": route_decision.get("Boundary_Policy"),
                **quality,
            }
        )
    return rows
