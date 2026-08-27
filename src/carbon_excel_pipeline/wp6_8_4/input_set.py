"""Stable fingerprint for one Current Run that may contain multiple input files."""

from __future__ import annotations

import hashlib
from typing import Any, Iterable


ROLE_PRIMARY = "主核算数据"
ROLE_ATTRIBUTE = "属性补充数据"
ROLE_LEDGER = "历史清册/因子参考"
ROLE_UNKNOWN = "未识别"


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def input_set_sha256(files: Iterable[dict[str, Any]]) -> str:
    """Hash every file SHA in a stable order. Router still uses the primary file SHA."""

    lines = []
    for item in files:
        digest = _text(item.get("sha256") or item.get("Input_SHA256")).upper()
        name = _text(item.get("name") or item.get("file_name"))
        if digest:
            lines.append(f"{digest}:{name}")
    if not lines:
        return ""
    payload = "\n".join(sorted(lines)).encode("utf-8")
    return hashlib.sha256(payload).hexdigest().upper()


def primary_file(files: Iterable[dict[str, Any]]) -> dict[str, Any] | None:
    items = list(files)
    for item in items:
        if _text(item.get("role")) == ROLE_PRIMARY:
            return item
    return items[0] if items else None


def attribute_files(files: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    return [item for item in files if _text(item.get("role")) == ROLE_ATTRIBUTE]


def ledger_files(files: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    return [item for item in files if _text(item.get("role")) == ROLE_LEDGER]
