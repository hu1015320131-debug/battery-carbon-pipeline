from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from carbon_excel_pipeline.ui.day9_controller import (
    latest_wp6_7_run,
    load_wp6_7_view,
)


class WP67UiTests(unittest.TestCase):
    def test_latest_run_and_view_only_load_backend_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            run = root / "WP6-7-TEST"
            run.mkdir()
            payloads = {
                "wp6_7_analysis_summary.json": {"run_id": "WP6-7-TEST", "status": "PASS"},
                "2024_data_quality_scorecard.json": {"Year": "2024", "metrics": []},
                "2025_data_quality_scorecard.json": {"Year": "2025", "metrics": []},
                "lineage_quality_summary.json": {"total": {"Complete_Lineage": 2}},
            }
            for name, payload in payloads.items():
                (run / name).write_text(json.dumps(payload), encoding="utf-8")
            for name in (
                "data_quality_issue_register.csv",
                "dimension_availability.csv",
                "2024_management_summary.csv",
                "2025_management_summary.csv",
                "2024_top_emission_contributors.csv",
                "2025_top_emission_contributors.csv",
                "2024_top_factor_impact.csv",
                "2025_top_factor_impact.csv",
            ):
                with (run / name).open("w", encoding="utf-8", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=["Record_ID"])
                    writer.writeheader()
                    writer.writerow({"Record_ID": "TEST-1"})
            self.assertEqual(latest_wp6_7_run(root), run)
            view = load_wp6_7_view(run)
            self.assertEqual(view["summary"]["status"], "PASS")
            self.assertEqual(view["scorecard_2024"]["Year"], "2024")
            self.assertEqual(view["contributors_2025"][0]["Record_ID"], "TEST-1")


if __name__ == "__main__":
    unittest.main()
