from __future__ import annotations

import json
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


class ProfileIsolationTests(unittest.TestCase):
    def test_only_public_profile_is_registered(self) -> None:
        registry = json.loads((ROOT / "config/profiles/profile_registry.json").read_text(encoding="utf-8"))
        self.assertEqual(registry["profiles"], ["public_synthetic_profile"])

    def test_public_profile_is_synthetic(self) -> None:
        public = json.loads(
            (ROOT / "config/profiles/public_synthetic_profile.json").read_text(encoding="utf-8")
        )
        self.assertEqual(public["classification"], "PUBLIC_SYNTHETIC_ONLY")
        self.assertFalse(public["formal_result_label_allowed"])
        self.assertEqual(public["record_id_regex"], "^SYN-[0-9]{6}$")


if __name__ == "__main__":
    unittest.main()
