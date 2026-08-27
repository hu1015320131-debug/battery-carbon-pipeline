"""Deterministic Record_ID namespaces for additional business units."""

from __future__ import annotations

import re
from copy import deepcopy
from typing import Any


CHINESE_NUMERALS = {
    "一": "1",
    "二": "2",
    "三": "3",
    "四": "4",
    "五": "5",
    "六": "6",
    "七": "7",
    "八": "8",
    "九": "9",
    "十": "10",
}


def namespace_for_unit(business_unit: str) -> str:
    text = (business_unit or "").strip()
    match = re.search(r"([一二三四五六七八九十])部", text)
    if match:
        return f"DY{CHINESE_NUMERALS[match.group(1)]}"
    match = re.search(r"(\d+)\s*部", text)
    if match:
        return f"DY{match.group(1)}"
    match = re.search(r"([一二三四五六七八九十])", text)
    if match:
        return f"DY{CHINESE_NUMERALS[match.group(1)]}"
    slug = re.sub(r"[^A-Za-z0-9]+", "", text).upper()[:8] or "EXT"
    return f"X{slug}"


def assign_additional_record_ids(
    records: list[dict[str, Any]],
    *,
    year: int,
    namespace: str,
) -> list[dict[str, Any]]:
    ordered = sorted(records, key=lambda item: int(item["Source_Row"]))
    output: list[dict[str, Any]] = []
    for index, record in enumerate(ordered, start=1):
        item = deepcopy(record)
        item["Record_ID"] = f"{year}-{namespace}-DX{index:06d}"
        output.append(item)
    return output
