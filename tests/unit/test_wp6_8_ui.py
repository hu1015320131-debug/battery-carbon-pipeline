from __future__ import annotations

import csv
import json
from pathlib import Path

from carbon_excel_pipeline.ui.day9_controller import (
    latest_wp6_8_run,
    load_wp6_8_view,
    wp6_8_download_artifacts,
)


def _write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def test_wp6_8_ui_loads_and_downloads_only_backend_artifacts(tmp_path: Path) -> None:
    run = tmp_path / "WP6-8-TEST"
    run.mkdir()
    summary = {
        "Run_ID": "WP6-8-TEST",
        "status": "PASS_WITH_WARNING",
        "record_counts": {"included": 1},
    }
    (run / "wp6_8_integration_summary.json").write_text(
        json.dumps(summary), encoding="utf-8"
    )
    (run / "run_summary.json").write_text(
        json.dumps({"Run_ID": "WP6-8-TEST"}), encoding="utf-8"
    )
    rows = [{"Record_ID": "TEST-1"}]
    for name in (
        "canonical_results.csv",
        "suggested_ledger_v1.csv",
        "audit_detail.csv",
        "excluded_records.csv",
        "dimension_availability.csv",
    ):
        _write_csv(run / name, rows)
    (run / "WP6_Result_Package.xlsx").write_bytes(b"synthetic-workbook")
    manifest = {
        "Run_ID": "WP6-8-TEST",
        "files": [
            {"File_Name": "audit_detail.csv"},
            {"File_Name": "WP6_Result_Package.xlsx"},
            {"File_Name": "../outside.txt"},
        ],
    }
    (run / "download_manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    (tmp_path / "outside.txt").write_text("outside", encoding="utf-8")

    assert latest_wp6_8_run(tmp_path) == run
    view = load_wp6_8_view(run)
    assert view["summary"]["status"] == "PASS_WITH_WARNING"
    assert view["audit"][0]["Record_ID"] == "TEST-1"
    artifacts = wp6_8_download_artifacts(run)
    assert {item["download_name"] for item in artifacts} == {
        "audit_detail.csv",
        "WP6_Result_Package.xlsx",
        "download_manifest.json",
    }


def test_streamlit_contains_no_wp6_8_business_recalculation() -> None:
    app_path = Path(__file__).resolve().parents[2] / "app" / "streamlit_app.py"
    source = app_path.read_text(encoding="utf-8")
    assert "run_wp6_8_integration" not in source
    assert "Activity_Data_kg / 1000" not in source
    assert "Emission_kgCO2e / 1000" not in source
    assert "build_scenario_analysis" not in source
    assert "build_data_quality_scorecard" not in source
