from __future__ import annotations

import json
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


class ContractRegistryTests(unittest.TestCase):
    def test_frozen_stage_counts(self) -> None:
        path = ROOT / "config/contracts/wp5_strict_contract.json"
        contract = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(
            contract["stage_field_counts"],
            {"D1": 45, "D2": 57, "D3": 36, "D4": 48, "D5": 56, "FROZEN_LINEAGE": 32},
        )

    def test_legacy_factor_is_explicitly_excluded(self) -> None:
        path = ROOT / "config/contracts/wp5_strict_contract.json"
        contract = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(contract["synthetic_factor"]["value"], "1.250000")
        self.assertEqual(
            contract["excluded_regression_values"]["policy"],
            "MUST_NOT_ENTER_FORMAL_EXPECTATIONS",
        )


if __name__ == "__main__":
    unittest.main()

