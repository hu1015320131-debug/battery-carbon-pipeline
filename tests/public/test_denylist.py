from __future__ import annotations

import unicodedata
from pathlib import Path

SKIP_DIR_NAMES = {".git", ".venv", "venv", "__pycache__", ".pytest_cache", ".mypy_cache"}
SKIP_SUFFIXES = {".png", ".jpg", ".xlsx", ".xls", ".pyc"}


def _tokens() -> list[str]:
    return [
        "Qi" + "Miao",
        "欣" + "旺达",
        "兰" + "溪",
        "锂" + "威",
        "L" + "WN",
        "2025-" + "LWN" + "-DX",
        "13.849" + "386",
        "14.655933372" + "055",
        "27.698" + "772",
        "69.246" + "930",
        "1384.93" + "86",
        "实习" + "课题",
        "C:\\" + "Users\\",
        "F1335E6F0D846337FC2964AB040CB4DCAFA56DDA1802E163A4E33A6CD0953FE9",
        "BE9F32289BE2725F051C23C868E40C69F23E01EB88901A1017FA3A3AD3E40C6D",
        "C882A4C4A189D6D09BDAF2B9C1941CBFFB70140C49E1134EA7CFDFA5AA8D726F",
        "58219EA46D2C16C6B769D0190C224A1EA150F63C896046BC78ACE1C401083AC2",
        "eco" + "invent 3.11",
        "电池" + "事业",
    ]


def test_public_tree_has_no_sensitive_tokens() -> None:
    root = Path(__file__).resolve().parents[2]
    self_name = Path(__file__).name
    hits: list[str] = []
    tokens = _tokens()
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if any(part in SKIP_DIR_NAMES for part in path.parts):
            continue
        if path.name == self_name:
            continue
        if path.suffix.lower() in SKIP_SUFFIXES:
            continue
        try:
            text = unicodedata.normalize("NFKC", path.read_text(encoding="utf-8"))
        except (UnicodeDecodeError, OSError):
            continue
        for token in tokens:
            if token in text:
                hits.append(f"{path.relative_to(root)}: {token}")
    assert hits == []
