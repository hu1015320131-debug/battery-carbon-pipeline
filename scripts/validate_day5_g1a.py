from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.activity.g1a_gate import evaluate_g1a


ROOT = Path(__file__).resolve().parents[1]


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return list(reader.fieldnames or []), list(reader)


def validate(
    run_dir: Path,
    *,
    standard_baseline: Path,
    activity_baseline: Path,
    third_party_baseline: Path,
) -> dict[str, Any]:
    standard_fields, standard = _read_csv(
        run_dir / "04_qc/day5_checked_standard_31_fields.csv"
    )
    activity_fields, activity = _read_csv(
        run_dir / "05_activity/day5_activity_36_fields.csv"
    )
    third_fields, third = _read_csv(
        run_dir / "06_third_party_input/day5_third_party_20_fields.csv"
    )
    _, record_items = _read_csv(
        run_dir / "05_activity/day5_record_open_items.csv"
    )
    _, interface_items = _read_csv(
        run_dir / "05_activity/day5_interface_open_items.csv"
    )
    quality_summary = json.loads(
        (run_dir / "04_qc/day5_quality_summary.json").read_text(encoding="utf-8")
    )
    quality_config = json.loads(
        (ROOT / "config/qc/day5_quality_rules.json").read_text(encoding="utf-8")
    )
    result = evaluate_g1a(
        standard_records=standard,
        standard_fields=standard_fields,
        activity_records=activity,
        activity_fields=activity_fields,
        third_party_records=third,
        third_party_fields=third_fields,
        quality_counts=quality_summary["status_counts"],
        record_open_item_count=len(record_items),
        interface_open_item_count=len(interface_items),
        standard_baseline_path=standard_baseline,
        activity_baseline_path=activity_baseline,
        third_party_baseline_path=third_party_baseline,
        quality_config=quality_config,
    )
    return {"validation_id": "DAY5_INDEPENDENT_G1A_VALIDATION_V1", **result}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-dir", required=True, type=Path)
    parser.add_argument("--standard-baseline", required=True, type=Path)
    parser.add_argument("--activity-baseline", required=True, type=Path)
    parser.add_argument("--third-party-baseline", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    result = validate(
        args.run_dir,
        standard_baseline=args.standard_baseline,
        activity_baseline=args.activity_baseline,
        third_party_baseline=args.third_party_baseline,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
