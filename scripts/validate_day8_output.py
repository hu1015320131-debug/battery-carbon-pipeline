"""Command-line wrapper for the independent Day 8 delivery validator."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from carbon_excel_pipeline.export.day8_validation import validate_day8_delivery


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a Day 8 G2 delivery without modifying it.")
    parser.add_argument("--report", required=True, type=Path)
    result = validate_day8_delivery(parser.parse_args().report)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
