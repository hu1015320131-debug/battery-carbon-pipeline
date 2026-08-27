"""Display-only numeric formatters. Never round values used in calculation."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any


def _decimal(value: Any) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    text = str(value).strip().replace(",", "")
    if not text or text in {"—", "-", "None", "nan", "NaN"}:
        return None
    for suffix in (" tCO2e", "tCO2e", " kg/year", "kg/year", " kgCO2e", "kgCO2e", "%"):
        if text.endswith(suffix):
            text = text[: -len(suffix)].strip()
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


def format_full_precision(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    number = _decimal(text)
    if number is None:
        return text
    return format(number, "f")


def _grouped(number: Decimal, places: str) -> str:
    quantized = number.quantize(Decimal(places), rounding=ROUND_HALF_UP)
    sign = "-" if quantized < 0 else ""
    quantized = abs(quantized)
    text = format(quantized, "f")
    if "." in text:
        integer, fraction = text.split(".", 1)
    else:
        integer, fraction = text, ""
    grouped = f"{int(integer):,}"
    if fraction:
        return f"{sign}{grouped}.{fraction}"
    return f"{sign}{grouped}"


def format_emission_display(value: Any, *, unit: str = "tCO2e") -> str:
    number = _decimal(value)
    if number is None:
        return "—"
    label = f"{_grouped(number, '0.01')} {unit}".strip()
    return label


def format_activity_display(value: Any, *, unit: str = "kg/year") -> str:
    number = _decimal(value)
    if number is None:
        return "—"
    return f"{_grouped(number, '0.01')} {unit}".strip()


def format_percentage_display(value: Any, *, places: int = 2) -> str:
    number = _decimal(value)
    if number is None:
        return "—"
    quant = "0." + ("0" * (places - 1)) + "1" if places > 0 else "1"
    return f"{_grouped(number, quant)}%"


def format_count_display(value: Any) -> str:
    number = _decimal(value)
    if number is None:
        return "—"
    return f"{int(number):,}"
