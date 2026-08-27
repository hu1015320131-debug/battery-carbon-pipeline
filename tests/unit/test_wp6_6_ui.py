from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from carbon_excel_pipeline.ui.day9_controller import (
    latest_wp6_6_run,
    load_wp6_6_view,
)


class WP66UiTests(unittest.TestCase):
    def test_latest_run_and_view_only_load_backend_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            run = root / "WP6-6-TEST"
            run.mkdir()
            payloads = {
                "wp6_6_analysis_summary.json": {"run_id": "WP6-6-TEST", "status": "PASS"},
                "scenario_abcd_summary.json": {"A": {"Emission_tCO2e": "1"}},
                "factor_effect_summary.json": {"Relative_EF_Change_Percent": "-5"},
                "activity_scope_effect_summary.json": {"status": "TEST"},
                "symmetric_decomposition.json": {"Reconciliation_Status": "PASS_EXACT"},
            }
            for name, payload in payloads.items():
                (run / name).write_text(json.dumps(payload), encoding="utf-8")
            for name in (
                "2024_factor_counterfactual.csv",
                "2025_factor_counterfactual.csv",
            ):
                with (run / name).open("w", encoding="utf-8", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=["Record_ID"])
                    writer.writeheader()
                    writer.writerow({"Record_ID": "TEST-1"})
            self.assertEqual(latest_wp6_6_run(root), run)
            view = load_wp6_6_view(run)
            self.assertEqual(view["summary"]["status"], "PASS")
            self.assertEqual(view["scenarios"]["A"]["Emission_tCO2e"], "1")
            self.assertEqual(view["records_2024"][0]["Record_ID"], "TEST-1")


if __name__ == "__main__":
    unittest.main()
