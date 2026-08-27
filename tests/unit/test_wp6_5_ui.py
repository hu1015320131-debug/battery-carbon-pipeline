from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from carbon_excel_pipeline.ui.day9_controller import (
    latest_wp6_5_run,
    load_wp6_5_view,
)


class WP65UiTests(unittest.TestCase):
    def test_latest_run_and_view_only_load_backend_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            run = root / "WP6-5-TEST"
            run.mkdir()
            (run / "independent_validation_summary.json").write_text(
                json.dumps({"run_id": "WP6-5-TEST", "status": "PASS"}),
                encoding="utf-8",
            )
            for name in (
                "2024_independent_validation.csv",
                "2025_independent_validation.csv",
                "independent_manual_samples.csv",
            ):
                with (run / name).open("w", encoding="utf-8", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=["Record_ID"])
                    writer.writeheader()
                    writer.writerow({"Record_ID": "TEST-1"})
            self.assertEqual(latest_wp6_5_run(root), run)
            view = load_wp6_5_view(run)
            self.assertEqual(view["summary"]["status"], "PASS")
            self.assertEqual(view["records_2024"][0]["Record_ID"], "TEST-1")


if __name__ == "__main__":
    unittest.main()
