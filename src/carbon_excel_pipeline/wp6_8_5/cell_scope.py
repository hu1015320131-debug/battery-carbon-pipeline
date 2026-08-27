"""Apply the cell business boundary before Capability and cell-specific QC."""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any


CELL_CATEGORY = "电芯"


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def is_cell_category(value: Any) -> bool:
    """Use the controlled first category segment; never infer cells from PCS alone."""

    return _text(value).split(".", 1)[0] == CELL_CATEGORY


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _sheet_business_unit(sheet_name: Any) -> str:
    text = _text(sheet_name)
    aliases = {
        "一部": "合成一部",
        "二部": "合成二部",
        "合成一部": "合成一部",
        "合成二部": "合成二部",
    }
    if text in aliases:
        return aliases[text]
    if "事业一部" in text:
        return "合成一部"
    if "事业二部" in text:
        return "合成二部"
    return ""


def _has_category_mapping(payload: dict[str, Any]) -> bool:
    return any(
        item.get("semantic_field") == "Purchase_Category"
        for item in payload.get("column_mappings") or []
    )


def _filter_payload(payload: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    output = dict(payload)
    records = list(payload.get("records") or [])
    if not _has_category_mapping(payload):
        return output, {
            "sheet_name": payload.get("sheet_name"),
            "raw_records": len(records),
            "cell_records": len(records),
            "non_cell_records": 0,
            "filter_applied": False,
            "reason": "PURCHASE_CATEGORY_NOT_MAPPED",
        }
    last_category = ""
    last_unit = ""
    sheet_unit = _sheet_business_unit(payload.get("sheet_name"))
    selected: list[dict[str, Any]] = []
    for record in records:
        item = dict(record)
        values = dict(item.get("values") or {})
        category = _text(values.get("Purchase_Category"))
        if category:
            last_category = category
        elif last_category:
            values["Purchase_Category"] = last_category
            category = last_category
        unit = _text(values.get("Business_Unit"))
        if unit:
            last_unit = unit
        elif last_unit:
            values["Business_Unit"] = last_unit
        elif sheet_unit:
            values["Business_Unit"] = sheet_unit
        item["values"] = values
        if is_cell_category(category):
            selected.append(item)
    output["records"] = selected
    output["record_count"] = len(selected)
    output["denominator_definition"] = (
        "Rows inside the controlled cell Purchase_Category boundary after recognition."
    )
    return output, {
        "sheet_name": payload.get("sheet_name"),
        "raw_records": len(records),
        "cell_records": len(selected),
        "non_cell_records": len(records) - len(selected),
        "filter_applied": True,
        "reason": "PURCHASE_CATEGORY_ROOT_EQUALS_CELL",
    }


def apply_cell_scope_to_run(run_dir: Path) -> dict[str, Any]:
    """Preserve raw recognition, then replace Capability input with cell-only rows."""

    import_dir = run_dir.expanduser().resolve() / "01_import"
    primary_path = import_dir / "recognized_records.json"
    sheets_path = import_dir / "recognized_records_by_sheet.json"
    if not primary_path.is_file():
        return {"status": "NOT_APPLIED", "reason": "RECOGNIZED_RECORDS_MISSING"}

    raw_primary = import_dir / "recognized_records_raw.json"
    raw_sheets = import_dir / "recognized_records_by_sheet_raw.json"
    if not raw_primary.is_file():
        shutil.copy2(primary_path, raw_primary)
    if sheets_path.is_file() and not raw_sheets.is_file():
        shutil.copy2(sheets_path, raw_sheets)

    primary_payload = _load_json(raw_primary)
    filtered_primary, primary_summary = _filter_payload(primary_payload)
    _write_json(primary_path, filtered_primary)

    sheet_summaries: list[dict[str, Any]] = [primary_summary]
    filtered_extra: list[dict[str, Any]] = []
    if raw_sheets.is_file():
        extra_payload = _load_json(raw_sheets)
        extra_sheets = extra_payload.get("sheets") or [] if isinstance(extra_payload, dict) else extra_payload
        for sheet in extra_sheets:
            filtered, summary = _filter_payload(sheet)
            sheet_summaries.append(summary)
            if filtered.get("records"):
                filtered_extra.append(filtered)
        if isinstance(extra_payload, dict):
            extra_payload = {**extra_payload, "sheets": filtered_extra}
        else:
            extra_payload = filtered_extra
        _write_json(sheets_path, extra_payload)

    raw_count = sum(int(item["raw_records"]) for item in sheet_summaries)
    cell_count = sum(int(item["cell_records"]) for item in sheet_summaries)
    summary = {
        "schema_version": "WP6_8_5_CELL_SCOPE_V1",
        "status": "PASS" if any(item["filter_applied"] for item in sheet_summaries) else "NOT_APPLIED",
        "boundary_rule": "Purchase_Category first segment exact equals 电芯",
        "raw_record_count": raw_count,
        "cell_record_count": cell_count,
        "non_cell_record_count": raw_count - cell_count,
        "capability_denominator": int(primary_summary["cell_records"]),
        "sheets": sheet_summaries,
        "raw_recognition_preserved": True,
    }
    _write_json(import_dir / "raw_data_statistics.json", {
        "raw_record_count": raw_count,
        "sheet_statistics": sheet_summaries,
        "purpose": "Recognition-range statistics only; non-cell rows do not enter cell Capability/QC.",
    })
    _write_json(import_dir / "cell_scope_summary.json", summary)
    return summary
