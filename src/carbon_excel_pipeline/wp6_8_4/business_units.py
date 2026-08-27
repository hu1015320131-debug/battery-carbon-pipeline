"""Detect and process additional business units without changing frozen 二部 baselines."""

from __future__ import annotations

import csv
import json
from decimal import Decimal
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.wp6_3.processing import (
    adapt_direct_mass,
    adapt_historical_ef,
    adapt_historical_ghg,
    calculate_and_validate,
    controlled_forward_fill,
    decimal_text,
    is_present,
)
from carbon_excel_pipeline.wp6_8_4.record_ids import (
    assign_additional_record_ids,
    namespace_for_unit,
)


CELL_ROOT = "电芯"
FROZEN_2024_UNIT = "合成二部"


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _category_root(value: Any) -> str:
    return _text(value).split(".", 1)[0]


def detect_business_units(records: list[dict[str, Any]], *, category_root: str = CELL_ROOT) -> list[str]:
    found: list[str] = []
    for item in records:
        values = item.get("values") or item
        unit = _text(values.get("Business_Unit") or item.get("Business_Unit"))
        if not unit or unit in found:
            continue
        if _category_root(values.get("Purchase_Category") or item.get("Purchase_Category")) == category_root:
            found.append(unit)
    return found


def unit_label_from_sheet(sheet_name: str) -> str:
    text = _text(sheet_name)
    if not text:
        return text
    if "事业" in text or text.startswith("合成"):
        return text
    if text.endswith("部") and len(text) <= 4:
        return f"合成{text}"
    return text


def _capability_map(run_dir: Path) -> dict[int, dict[str, str]]:
    rows = _read_csv(run_dir / "02_capability" / "record_capabilities.csv")
    return {int(row["Source_Row"]): row for row in rows if row.get("Source_Row")}


def additional_direct_mass_rows(
    run_dir: Path,
    *,
    existing_canonical: list[dict[str, Any]],
    policy: dict[str, Any],
    run_id: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    records_path = run_dir / "01_import" / "recognized_records.json"
    if not records_path.is_file():
        return [], [], detect_business_units(existing_canonical)
    payload = _load_json(records_path)
    filled = controlled_forward_fill(
        payload.get("records") or [],
        fields=tuple(policy.get("controlled_forward_fill_fields") or ("Business_Unit", "Purchase_Type", "Purchase_Category")),
    )
    detected = detect_business_units(filled)
    frozen_unit = str((policy.get("boundary") or {}).get("business_unit_exact") or FROZEN_2024_UNIT)
    used_rows = {int(row.get("Source_Row") or 0) for row in existing_canonical}
    units = payload.get("units") or {}
    capabilities = _capability_map(run_dir)
    year = int(policy.get("year") or 2024)
    threshold = Decimal(str(policy.get("historical_difference_reporting_threshold_tco2e") or "0.000000001"))
    extra: list[dict[str, Any]] = []
    historical: list[dict[str, Any]] = []
    fingerprint = str(payload.get("input_fingerprint") or "")
    for unit in detected:
        if unit == frozen_unit:
            continue
        selected = [
            item
            for item in filled
            if _text((item.get("values") or {}).get("Business_Unit")) == unit
            and _category_root((item.get("values") or {}).get("Purchase_Category")) == CELL_ROOT
            and int(item["Source_Row"]) not in used_rows
        ]
        assigned = assign_additional_record_ids(
            selected, year=year, namespace=namespace_for_unit(unit)
        )
        for record in assigned:
            values = record.get("values") or {}
            row_number = int(record["Source_Row"])
            cap = capabilities.get(row_number, {})
            supported = [item for item in str(cap.get("Supported_Activity_Paths") or "").split("|") if item]
            warnings: list[str] = []
            blocking: list[str] = []
            if supported and "DIRECT_REPORTED_MASS" not in supported:
                blocking.append("REQUESTED_ACTIVITY_PATH_NOT_SUPPORTED")
            activity = adapt_direct_mass(
                values.get("Reported_Activity_Value"),
                units.get("Reported_Activity_Value"),
            )
            factor = adapt_historical_ef(
                values.get("EF_Value"), units.get("EF_Value"), values.get("EF_Source")
            )
            historical_ghg = adapt_historical_ghg(
                values.get("Historical_GHG_Value"), units.get("Historical_GHG_Value")
            )
            warnings.extend(activity["warnings"])
            warnings.extend(factor["warnings"])
            warnings.extend(historical_ghg["warnings"])
            blocking.extend(activity["blocking"])
            blocking.extend(factor["blocking"])
            calculation = None
            if not blocking:
                calculation = calculate_and_validate(
                    activity_kg=activity["activity_kg"],
                    ef_value=factor["normalized_value"],
                    historical_tco2e=historical_ghg["normalized_tco2e"],
                    difference_threshold_tco2e=threshold,
                )
            warnings = sorted(set(warnings))
            blocking = sorted(set(blocking))
            qc_status = "BLOCKED" if blocking else ("WARNING" if warnings else "PASS")
            extra.append(
                {
                    "Record_ID": record["Record_ID"],
                    "Year": str(year),
                    "Source_File": payload.get("workbook_name"),
                    "Source_SHA256": fingerprint,
                    "Source_Sheet": payload.get("sheet_name"),
                    "Source_Row": row_number,
                    "Business_Unit": values.get("Business_Unit"),
                    "Purchase_Type": values.get("Purchase_Type"),
                    "Purchase_Category": values.get("Purchase_Category"),
                    "Product_Description": values.get("Product_Description"),
                    "Original_Activity_Value": values.get("Reported_Activity_Value"),
                    "Original_Activity_Unit": units.get("Reported_Activity_Value"),
                    "Activity_Data_kg": decimal_text(activity["activity_kg"]),
                    "Activity_Method": "DIRECT_REPORTED_MASS",
                    "Unit_Conversion_Factor": decimal_text(activity["conversion_factor"]),
                    "EF_Value": decimal_text(factor["normalized_value"]),
                    "EF_Unit": "kgCO2e/kg" if factor["normalized_value"] is not None else units.get("EF_Value"),
                    "EF_Source": values.get("EF_Source"),
                    "EF_Usage": "SOURCE_EMBEDDED_FACTOR",
                    "Emission_kgCO2e": decimal_text(calculation["emission_kg"] if calculation else None),
                    "Emission_tCO2e": decimal_text(calculation["emission_t"] if calculation else None),
                    "Historical_GHG_Value": decimal_text(historical_ghg["normalized_tco2e"]),
                    "Historical_GHG_Unit": "tCO2e/year" if historical_ghg["normalized_tco2e"] is not None else "",
                    "Historical_Difference": decimal_text(calculation["difference_t"] if calculation else None),
                    "QC_Status": qc_status,
                    "Calculation_QC": "BLOCKED" if blocking else "PASS",
                    "Governance_QC": "WARNING" if warnings else "PASS",
                    "Warning_Codes": "|".join(warnings),
                    "Blocking_Codes": "|".join(blocking),
                    "Simulation_Flag": "TRUE",
                    "Production_Eligible": "FALSE",
                    "Run_ID": run_id,
                    "Boundary_Policy": "DETECTED_BUSINESS_UNIT_CELL",
                    "Factor_Route": "SOURCE_EMBEDDED_FACTOR" if factor["normalized_value"] is not None else "FACTOR_NOT_AVAILABLE",
                }
            )
            if calculation:
                historical.append(
                    {
                        "Record_ID": record["Record_ID"],
                        "Source_Row": row_number,
                        "Calculated_Emission_tCO2e": decimal_text(calculation["emission_t"]),
                        "Historical_Emission_tCO2e": decimal_text(historical_ghg["normalized_tco2e"]),
                        "Difference_tCO2e": decimal_text(calculation["difference_t"]),
                        "Difference_Percent": decimal_text(calculation.get("difference_percent")),
                        "Validation_Status": calculation["validation_status"],
                    }
                )
    detected_all = list(dict.fromkeys(
        detect_business_units(existing_canonical) + detected
    ))
    return extra, historical, detected_all


def additional_pcs_sheet_rows(
    run_dir: Path,
    *,
    existing_canonical: list[dict[str, Any]],
    run_id: str,
    input_sha256: str,
    source_name: str,
) -> tuple[list[dict[str, Any]], list[str]]:
    primary_path = run_dir / "01_import" / "recognized_records.json"
    extra_path = run_dir / "01_import" / "recognized_records_by_sheet.json"
    if not primary_path.is_file() and not extra_path.is_file():
        return [], detect_business_units(existing_canonical)
    sheets: list[dict[str, Any]] = []
    if primary_path.is_file():
        sheets.append(_load_json(primary_path))
    if extra_path.is_file():
        extra = _load_json(extra_path)
        sheets.extend(extra.get("sheets") or [] if isinstance(extra, dict) else extra)
    existing_keys = {
        (_text(row.get("Source_Sheet")), int(row.get("Source_Row") or 0))
        for row in existing_canonical
    }
    grouped: dict[str, list[tuple[dict[str, Any], dict[str, Any]]]] = {}
    for payload in sheets:
        sheet_name = _text(payload.get("sheet_name"))
        if not sheet_name:
            continue
        default_unit = unit_label_from_sheet(sheet_name)
        for raw in payload.get("records") or []:
            values = raw.get("values") or {}
            category = values.get("Purchase_Category")
            if _category_root(category) != CELL_ROOT:
                continue
            key = (sheet_name, int(raw.get("Source_Row") or 0))
            if key in existing_keys:
                continue
            unit = _text(values.get("Business_Unit")) or default_unit
            grouped.setdefault(unit, []).append((payload, raw))
    extra: list[dict[str, Any]] = []
    for unit_label, scoped in grouped.items():
        scoped.sort(key=lambda item: (_text(item[0].get("sheet_name")), int(item[1].get("Source_Row") or 0)))
        year = 2025
        assigned = assign_additional_record_ids(
            [
                {"Source_Row": raw.get("Source_Row"), "values": raw.get("values") or {}}
                for _, raw in scoped
            ],
            year=year,
            namespace=namespace_for_unit(unit_label),
        )
        for record, (payload, raw) in zip(assigned, scoped, strict=True):
            sheet_name = _text(payload.get("sheet_name"))
            units = payload.get("units") or {}
            values = record.get("values") or raw.get("values") or {}
            category = values.get("Purchase_Category")
            pcs = values.get("Quantity_PCS")
            weight = values.get("Unit_Weight")
            activity = None
            blocking: list[str] = []
            try:
                pcs_number = Decimal(str(pcs))
                weight_number = Decimal(str(weight))
                if pcs_number <= 0 or weight_number <= 0:
                    blocking.append("PCS_OR_WEIGHT_INVALID")
                else:
                    weight_unit = units.get("Unit_Weight") or "g/PCS"
                    factor = Decimal("0.001") if weight_unit in {"g/PCS", "g"} else Decimal("1")
                    activity = pcs_number * weight_number * factor
            except Exception:
                blocking.append("PCS_OR_WEIGHT_INVALID")
            extra.append(
                {
                    "Record_ID": record["Record_ID"],
                    "Year": str(year),
                    "Source_File": payload.get("workbook_name") or source_name,
                    "Source_SHA256": input_sha256,
                    "Source_Sheet": sheet_name,
                    "Source_Row": record["Source_Row"],
                    "Business_Unit": unit_label,
                    "Purchase_Category": category,
                    "Product_Description": values.get("Product_Description"),
                    "Quantity_PCS": pcs,
                    "Unit_Weight": weight,
                    "Unit_Weight_Unit": units.get("Unit_Weight"),
                    "Original_Activity_Value": values.get("Reported_Purchase_Quantity") or pcs,
                    "Activity_Data_kg": decimal_text(activity) if activity is not None and not blocking else "",
                    "Activity_Unit": "kg/year",
                    "Activity_Method": "PCS_WEIGHT_DERIVED",
                    "EF_Value": "",
                    "EF_Source": "",
                    "Emission_kgCO2e": "",
                    "Emission_tCO2e": "",
                    "QC_Status": "BLOCKED" if blocking else "WARNING",
                    "Calculation_QC": "BLOCKED",
                    "Governance_QC": "WARNING",
                    "Warning_Codes": "",
                    "Blocking_Codes": "|".join(blocking + ["FACTOR_NOT_AVAILABLE", "BOUNDARY_POLICY_NOT_AVAILABLE"]),
                    "Simulation_Flag": "TRUE",
                    "Production_Eligible": "FALSE",
                    "Run_ID": run_id,
                    "Boundary_Policy": "DETECTED_SHEET_NOT_SHA_GATED",
                    "Factor_Route": "FACTOR_NOT_AVAILABLE",
                    "Overall_Status": "PARTIAL_RESULT",
                }
            )
    detected = list(dict.fromkeys(
        detect_business_units(existing_canonical)
        + detect_business_units(extra)
        + list(grouped)
    ))
    return extra, [item for item in detected if item]


def synthetic_supplier_subset(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        row
        for row in rows
        if _text(row.get("Record_ID")).startswith("2024-DY2-SYNA-DX")
        or _text(row.get("Record_ID")).startswith("2025-DY2-SYNA-DX")
    ]


def filter_by_business_unit(
    rows: list[dict[str, Any]],
    business_unit: str | None,
) -> list[dict[str, Any]]:
    if not business_unit or business_unit in {"全部", "ALL"}:
        return list(rows)
    wanted = _text(business_unit)
    return [row for row in rows if _text(row.get("Business_Unit")) == wanted]
