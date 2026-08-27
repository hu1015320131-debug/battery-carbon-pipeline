"""Generate RID_V2 once and propagate it through the current Run."""

from __future__ import annotations

import csv
import json
import re
from collections import Counter, defaultdict
from copy import deepcopy
from pathlib import Path
from typing import Any, Mapping


RECORD_ID_SCHEMA_VERSION = "RID_V2"
LEGACY_FIELDS = {"Legacy_Record_ID", "Old_Record_ID"}
LEGACY_ID_PATTERN = re.compile(r"^\d{4}-(?:SYNA-DX|DY\d+-DX)\d{6}$")
RECORD_ID_PATTERN = re.compile(
    r"^(?P<year>\d{4})-(?P<business_unit>DY[1-9]\d*)-"
    r"(?P<supplier>[A-Z0-9]+)-(?P<material>DX|PCB|other)(?P<serial>\d{6})$"
)


class RecordIDSchemaError(ValueError):
    """Block the Run when RID_V2 cannot be generated deterministically."""


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def load_record_id_config(path: Path) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != RECORD_ID_SCHEMA_VERSION:
        raise RecordIDSchemaError("Record_ID 配置不是 RID_V2。")
    return payload


def business_unit_code(value: Any, config: Mapping[str, Any]) -> str:
    text = _text(value)
    controlled = config.get("business_unit_codes") or {}
    if text in controlled:
        return str(controlled[text])
    chinese = {
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
        "十": 10,
    }
    match = re.search(r"事业([一二三四五六七八九十]|\d+)部", text)
    if match:
        token = match.group(1)
        number = chinese.get(token, int(token) if token.isdigit() else 0)
        if number:
            return f"DY{number}"
    raise RecordIDSchemaError(f"Business_Unit 没有受控事业部代码：{text or '<空>'}")


def _supplier_aliases(config: Mapping[str, Any]) -> dict[str, str]:
    aliases: dict[str, str] = {}
    for code, values in (config.get("supplier_codes") or {}).items():
        controlled_code = _text(code).upper()
        if not re.fullmatch(r"[A-Z0-9]+", controlled_code):
            raise RecordIDSchemaError(f"供应商代码不合法：{code}")
        aliases[controlled_code.casefold()] = controlled_code
        for value in values or []:
            aliases[_text(value).casefold()] = controlled_code
    return aliases


def supplier_code(row: Mapping[str, Any], config: Mapping[str, Any]) -> str:
    aliases = _supplier_aliases(config)
    for field in (
        "Supplier_Code",
        "Supplier_ID",
        "Supplier_Abbreviation",
        "Supplier",
        "Supplier_Name",
    ):
        value = _text(row.get(field))
        if value and value.casefold() in aliases:
            return aliases[value.casefold()]
    return _text(config.get("unknown_supplier_code") or "UNK").upper()


def material_code(row: Mapping[str, Any], config: Mapping[str, Any]) -> str:
    mappings = config.get("material_codes") or {}
    explicit = _text(row.get("Material_Code") or row.get("Activity_Category_Code"))
    if explicit in mappings.values():
        return explicit
    category = _text(row.get("Purchase_Category") or row.get("Purchase_Type"))
    root = category.split(".", 1)[0]
    if root in mappings:
        return str(mappings[root])
    raise RecordIDSchemaError(f"物料类型没有受控 Material Code：{root or '<空>'}")


def _integer(value: Any, *, default: int) -> int:
    text = _text(value)
    try:
        return int(text)
    except (TypeError, ValueError):
        return default


def _source_key(row: Mapping[str, Any], original_index: int) -> tuple[Any, ...]:
    return (
        _integer(row.get("Source_File_Order"), default=0),
        _integer(row.get("Source_Sheet_Order"), default=0),
        _text(row.get("Source_File")),
        _text(row.get("Source_Sheet")),
        _integer(row.get("Source_Row"), default=10**12),
        original_index,
    )


def validate_record_id(record_id: Any) -> bool:
    return RECORD_ID_PATTERN.fullmatch(_text(record_id)) is not None


def assign_record_ids(
    records: list[dict[str, Any]],
    *,
    config: Mapping[str, Any],
    source_sheet_orders: Mapping[str, int] | None = None,
) -> tuple[list[dict[str, Any]], dict[str, str], dict[str, Any]]:
    """Assign one stable RID_V2 per row after Supplier enrichment."""

    output = [deepcopy(row) for row in records]
    sheet_orders = source_sheet_orders or {}
    grouped: dict[tuple[str, str, str, str], list[tuple[int, dict[str, Any]]]] = defaultdict(list)
    source_keys: list[tuple[str, str, str]] = []
    for index, row in enumerate(output):
        for field in LEGACY_FIELDS:
            if field in row:
                raise RecordIDSchemaError(f"RID_V2 不允许字段 {field}。")
        year = _text(row.get("Year"))
        if not re.fullmatch(r"\d{4}", year):
            raise RecordIDSchemaError(f"Year 必须来自记录且为 4 位：{year or '<空>'}")
        unit = business_unit_code(row.get("Business_Unit"), config)
        supplier = supplier_code(row, config)
        material = material_code(row, config)
        if "Source_Sheet_Order" not in row:
            row["Source_Sheet_Order"] = sheet_orders.get(_text(row.get("Source_Sheet")), 0)
        source_identity = (
            _text(row.get("Source_File")),
            _text(row.get("Source_Sheet")),
            _text(row.get("Source_Row")),
        )
        source_keys.append(source_identity)
        grouped[(year, unit, supplier, material)].append((index, row))

    duplicates = [key for key, count in Counter(source_keys).items() if count > 1]
    if duplicates:
        raise RecordIDSchemaError(f"稳定物理源键重复，无法生成唯一 Record_ID：{duplicates[0]}")

    id_mapping: dict[str, str] = {}
    scope_counts: dict[str, int] = {}
    for scope, members in grouped.items():
        ordered = sorted(members, key=lambda item: _source_key(item[1], item[0]))
        year, unit, supplier, material = scope
        scope_name = "|".join(scope)
        scope_counts[scope_name] = len(ordered)
        for serial, (_, row) in enumerate(ordered, start=1):
            old_id = _text(row.get("Record_ID"))
            new_id = f"{year}-{unit}-{supplier}-{material}{serial:06d}"
            if not validate_record_id(new_id):
                raise RecordIDSchemaError(f"生成的 Record_ID 未通过 RID_V2 校验：{new_id}")
            if old_id:
                if old_id in id_mapping and id_mapping[old_id] != new_id:
                    raise RecordIDSchemaError(f"旧 Record_ID 重复映射：{old_id}")
                id_mapping[old_id] = new_id
            row["Record_ID"] = new_id
            row["Record_ID_Schema_Version"] = RECORD_ID_SCHEMA_VERSION

    ids = [_text(row.get("Record_ID")) for row in output]
    if not ids or any(not value for value in ids):
        raise RecordIDSchemaError("Record_ID 必须 100% 非空。")
    if len(ids) != len(set(ids)):
        raise RecordIDSchemaError("Record_ID 必须 100% 唯一。")
    if any(not validate_record_id(value) for value in ids):
        raise RecordIDSchemaError("存在未通过 RID_V2 格式校验的 Record_ID。")
    if any(LEGACY_ID_PATTERN.fullmatch(value) for value in ids):
        raise RecordIDSchemaError("新 Current Run 中仍存在旧格式 Record_ID。")

    summary = {
        "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
        "status": "PASS",
        "record_count": len(ids),
        "unique_count": len(set(ids)),
        "missing_count": 0,
        "scope_counts": scope_counts,
        "stable_order": "Source_File_Order -> Source_Sheet_Order -> Source_Row",
        "legacy_fields_present": False,
        "legacy_format_present": False,
    }
    return output, id_mapping, summary


def remap_record_ids_in_rows(
    rows: list[dict[str, Any]], id_mapping: Mapping[str, str]
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for row in rows:
        item = deepcopy(row)
        for key, value in list(item.items()):
            if _text(value) in id_mapping:
                item[key] = id_mapping[_text(value)]
        if item.get("Record_ID"):
            item["Record_ID_Schema_Version"] = RECORD_ID_SCHEMA_VERSION
        output.append(item)
    return output


def _replace_text(value: str, id_mapping: Mapping[str, str]) -> tuple[str, int]:
    rewritten = value
    replacements = 0
    for old_id, new_id in id_mapping.items():
        count = rewritten.count(old_id)
        if count:
            rewritten = rewritten.replace(old_id, new_id)
            replacements += count
    return rewritten, replacements


def _remap_json(value: Any, id_mapping: Mapping[str, str]) -> Any:
    if isinstance(value, dict):
        output: dict[str, Any] = {}
        for key, item in value.items():
            rewritten_key, _ = _replace_text(str(key), id_mapping)
            output[rewritten_key] = _remap_json(item, id_mapping)
        return output
    if isinstance(value, list):
        return [_remap_json(item, id_mapping) for item in value]
    if isinstance(value, str):
        return _replace_text(value, id_mapping)[0]
    return value


def propagate_record_ids_in_run(run_dir: Path, id_mapping: Mapping[str, str]) -> dict[str, Any]:
    """Replace exact provisional IDs in current-Run CSV/JSON/Markdown artifacts."""

    run = run_dir.expanduser().resolve()
    changed_files: list[str] = []
    replacements = 0
    for path in sorted(run.rglob("*")):
        if not path.is_file() or "00_input_copy" in path.parts:
            continue
        suffix = path.suffix.lower()
        if suffix == ".csv":
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.reader(handle))
            changed = False
            for row in rows:
                for index, value in enumerate(row):
                    rewritten, count = _replace_text(value, id_mapping)
                    if count:
                        row[index] = rewritten
                        replacements += count
                        changed = True
            if changed:
                with path.open("w", encoding="utf-8-sig", newline="") as handle:
                    csv.writer(handle, lineterminator="\n").writerows(rows)
                changed_files.append(str(path.relative_to(run)))
        elif suffix == ".json":
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except (UnicodeDecodeError, json.JSONDecodeError):
                continue
            rewritten = _remap_json(payload, id_mapping)
            if rewritten != payload:
                before = json.dumps(payload, ensure_ascii=False)
                _, count = _replace_text(before, id_mapping)
                replacements += count
                path.write_text(
                    json.dumps(rewritten, ensure_ascii=False, indent=2) + "\n",
                    encoding="utf-8",
                )
                changed_files.append(str(path.relative_to(run)))
        elif suffix == ".md":
            text = path.read_text(encoding="utf-8")
            rewritten = text
            rewritten, count = _replace_text(text, id_mapping)
            replacements += count
            if rewritten != text:
                path.write_text(rewritten, encoding="utf-8")
                changed_files.append(str(path.relative_to(run)))
    return {
        "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
        "status": "PASS",
        "mapped_record_count": len(id_mapping),
        "replacement_count": replacements,
        "changed_file_count": len(changed_files),
        "changed_files": changed_files,
    }
