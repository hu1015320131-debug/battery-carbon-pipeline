import csv
import json
from pathlib import Path

import pytest

from carbon_excel_pipeline.wp6_8_6.record_ids import (
    RecordIDSchemaError,
    assign_record_ids,
    load_record_id_config,
    propagate_record_ids_in_run,
    validate_record_id,
)


ROOT = Path(__file__).resolve().parents[2]
CONFIG = load_record_id_config(ROOT / "config" / "wp6" / "record_id_schema_v2.json")


def _row(
    source_row: int,
    *,
    year: str = "2025",
    unit: str = "合成一部",
    supplier: str = "SYNB",
    category: str = "电芯",
) -> dict:
    return {
        "Record_ID": f"TEMP-{source_row}",
        "Year": year,
        "Business_Unit": unit,
        "Supplier": supplier,
        "Purchase_Category": category,
        "Source_File": "main.xlsx",
        "Source_Sheet": "Sheet1",
        "Source_Row": source_row,
    }


@pytest.mark.parametrize(
    ("row", "expected"),
    [
        (_row(2), "2025-DY1-SYNB-DX000001"),
        (_row(2, supplier=""), "2025-DY1-UNK-DX000001"),
        (
            _row(2, year="2024", unit="合成二部", supplier="SYNA", category="PCB"),
            "2024-DY2-SYNA-PCB000001",
        ),
        (
            _row(2, unit="合成二部", supplier="SYNA", category="其他"),
            "2025-DY2-SYNA-other000001",
        ),
    ],
)
def test_standard_controlled_record_ids(row: dict, expected: str):
    output, _, summary = assign_record_ids([row], config=CONFIG)
    assert output[0]["Record_ID"] == expected
    assert output[0]["Record_ID_Schema_Version"] == "RID_V2"
    assert summary["status"] == "PASS"
    assert validate_record_id(expected)


def test_serials_are_independent_by_supplier_and_stable_by_source_row():
    rows = [_row(20), _row(5, supplier="SYNA"), _row(10), _row(6, supplier="SYNA")]
    first, _, _ = assign_record_ids(rows, config=CONFIG)
    second, _, _ = assign_record_ids(list(reversed(rows)), config=CONFIG)
    by_source_first = {row["Source_Row"]: row["Record_ID"] for row in first}
    by_source_second = {row["Source_Row"]: row["Record_ID"] for row in second}
    assert by_source_first == by_source_second == {
        5: "2025-DY1-SYNA-DX000001",
        6: "2025-DY1-SYNA-DX000002",
        10: "2025-DY1-SYNB-DX000001",
        20: "2025-DY1-SYNB-DX000002",
    }


def test_unknown_chinese_supplier_is_not_used_to_invent_a_code():
    output, _, _ = assign_record_ids([_row(2, supplier="未登记供应商")], config=CONFIG)
    assert output[0]["Record_ID"] == "2025-DY1-UNK-DX000001"


def test_invalid_material_and_legacy_fields_block_instead_of_guessing():
    with pytest.raises(RecordIDSchemaError, match="Material Code"):
        assign_record_ids([_row(2, category="结构件")], config=CONFIG)
    legacy = _row(2)
    legacy["Legacy_Record_ID"] = "2025-DY2-SYNA-DX000001"
    with pytest.raises(RecordIDSchemaError, match="Legacy_Record_ID"):
        assign_record_ids([legacy], config=CONFIG)


def test_duplicate_physical_source_key_blocks():
    with pytest.raises(RecordIDSchemaError, match="物理源键重复"):
        assign_record_ids([_row(2), _row(2, supplier="SYNA")], config=CONFIG)


def test_propagation_replaces_exact_ids_in_csv_and_json(tmp_path: Path):
    run = tmp_path / "RUN"
    csv_path = run / "05_activity" / "activity.csv"
    csv_path.parent.mkdir(parents=True)
    with csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["Record_ID", "Activity"])
        writer.writeheader()
        writer.writerow({"Record_ID": "TEMP-2", "Activity": "TRACE:TEMP-2|10"})
    json_path = run / "09_calculation" / "audit.json"
    json_path.parent.mkdir(parents=True)
    json_path.write_text(
        json.dumps({"by_id": {"TEMP-2": {"Record_ID": "TEMP-2", "Emission": "20"}}}),
        encoding="utf-8",
    )
    summary = propagate_record_ids_in_run(
        run, {"TEMP-2": "2025-DY1-SYNB-DX000001"}
    )
    assert summary["changed_file_count"] == 2
    assert "TEMP-2" not in csv_path.read_text(encoding="utf-8-sig")
    assert "TEMP-2" not in json_path.read_text(encoding="utf-8")
    assert "2025-DY1-SYNB-DX000001" in csv_path.read_text(encoding="utf-8-sig")
