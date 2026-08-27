from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from carbon_excel_pipeline.ui.day9_controller import (
    latest_wp6_4_run,
    load_wp6_4_view,
)


class WP64UiTests(unittest.TestCase):
    def test_latest_formal_run_and_view_use_backend_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            run = root / "WP6-4-TEST"
            run.mkdir()
            (run / "wp6_4_summary.json").write_text(
                json.dumps({"run_id": "WP6-4-TEST", "status": "PASS"}),
                encoding="utf-8",
            )
            for name in (
                "2025_strict_regression.csv",
                "2025_qc_regression.csv",
                "2025_scope_comparison.csv",
            ):
                with (run / name).open("w", encoding="utf-8", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=["Record_ID"])
                    writer.writeheader()
                    writer.writerow({"Record_ID": "2025-DY2-SYNA-DX000001"})
            (run / "2024_shared_model_forward_compatibility.json").write_text(
                json.dumps({"controlled_forward_fill": {}}), encoding="utf-8"
            )
            self.assertEqual(latest_wp6_4_run(root), run)
            view = load_wp6_4_view(run)
            self.assertEqual(view["summary"]["status"], "PASS")
            self.assertEqual(view["strict_records"][0]["Record_ID"], "2025-DY2-SYNA-DX000001")


if __name__ == "__main__":
    unittest.main()
