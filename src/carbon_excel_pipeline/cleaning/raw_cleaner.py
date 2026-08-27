"""Day 3 raw-value preservation and strict cleaning primitives."""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any


WHITESPACE_PATTERN = re.compile(r"\s+")
PROJECT_PREFIX_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9-]{1,29}$")
CELL_MODEL_PATTERN = re.compile(r"(?<!\d)(\d{6})(?!\d)")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\u3000", " ").strip()
    return WHITESPACE_PATTERN.sub(" ", text)


def contains_marker(text: str, marker: str) -> bool:
    return marker.casefold() in text.casefold()


def extract_project_prefix_candidate(text: str) -> str:
    for token in text.split(" "):
        stripped = token.strip(";,，；:：()（）[]【】")
        if PROJECT_PREFIX_PATTERN.fullmatch(stripped):
            return stripped
    return ""


def extract_cell_model_candidate(text: str) -> str:
    match = CELL_MODEL_PATTERN.search(text)
    return match.group(1) if match else ""


def canonical_decimal(value: Decimal) -> str:
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


@dataclass(frozen=True, slots=True)
class NumericCleanResult:
    cleaned_value: str
    issue_code: str

    @property
    def is_valid(self) -> bool:
        return not self.issue_code


def parse_strict_positive_decimal(
    value: Any,
    *,
    field_code: str,
    integer_required: bool = False,
) -> NumericCleanResult:
    if value is None or value == "":
        return NumericCleanResult("", f"{field_code}_BLANK")
    if isinstance(value, bool):
        return NumericCleanResult("", f"{field_code}_BOOLEAN_NOT_ALLOWED")
    if isinstance(value, str):
        return NumericCleanResult("", f"{field_code}_TEXT_NOT_ALLOWED")
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return NumericCleanResult("", f"{field_code}_NON_FINITE")
    try:
        parsed = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return NumericCleanResult("", f"{field_code}_INVALID_NUMBER")
    if not parsed.is_finite():
        return NumericCleanResult("", f"{field_code}_NON_FINITE")
    if parsed <= 0:
        return NumericCleanResult("", f"{field_code}_NOT_STRICTLY_POSITIVE")
    if integer_required and parsed != parsed.to_integral_value():
        return NumericCleanResult("", f"{field_code}_NOT_INTEGER")
    return NumericCleanResult(canonical_decimal(parsed), "")


@dataclass(frozen=True, slots=True)
class UnitCleanResult:
    cleaned_unit: str
    issue_code: str

    @property
    def is_valid(self) -> bool:
        return not self.issue_code


def map_unit_strict(
    raw_unit: Any,
    *,
    field_code: str,
    mappings: dict[str, str],
) -> UnitCleanResult:
    if raw_unit is None or raw_unit == "":
        return UnitCleanResult("", f"{field_code}_UNIT_BLANK")
    if not isinstance(raw_unit, str):
        return UnitCleanResult("", f"{field_code}_UNIT_NOT_TEXT")
    mapped = mappings.get(raw_unit)
    if mapped is None:
        return UnitCleanResult("", f"{field_code}_UNIT_UNMAPPED_EXACT")
    return UnitCleanResult(mapped, "")

