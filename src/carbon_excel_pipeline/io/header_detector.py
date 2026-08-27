"""Compatibility facade for the WP6-1 recognition core."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from carbon_excel_pipeline.io.recognition import legacy_header_result, recognize_sheet
from carbon_excel_pipeline.io.semantic_registry import (
    SemanticFieldRegistry,
    normalize_header,
    parse_header_unit,
    semantic_header_key,
)


def load_alias_config(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
    required = {"scan_rows", "fields"}
    missing = sorted(required - config.keys())
    if missing:
        raise ValueError(f"Header alias configuration is missing keys: {missing}")
    SemanticFieldRegistry(config)
    return config


def detect_header(worksheet: Worksheet, config: dict[str, Any]) -> dict[str, Any]:
    """Return the frozen Day 2 shape backed by WP6-1 semantic recognition."""
    registry = SemanticFieldRegistry(config)
    result = recognize_sheet(worksheet, sheet_index=0, registry=registry)
    return legacy_header_result(result, registry=registry)


__all__ = [
    "detect_header",
    "load_alias_config",
    "normalize_header",
    "parse_header_unit",
    "semantic_header_key",
]
