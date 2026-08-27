from __future__ import annotations

import unittest

from carbon_excel_pipeline.bootstrap import run_startup_checks


class StartupCheckTests(unittest.TestCase):
    def test_scaffold_reaches_g0_candidate(self) -> None:
        result = run_startup_checks()
        self.assertEqual(result["status"], "PASS", result["errors"])
        self.assertEqual(result["gate_candidate"], "G0_READY_TO_BUILD")
        self.assertFalse(result["business_pipeline_executed"])
        self.assertFalse(result["real_data_read"])


if __name__ == "__main__":
    unittest.main()

