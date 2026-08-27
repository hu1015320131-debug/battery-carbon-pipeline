from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


class PublicBoundaryTests(unittest.TestCase):
    def test_public_profile_is_synthetic_and_not_a_formal_result(self) -> None:
        text = (ROOT / "config/profiles/public_synthetic_profile.json").read_text(encoding="utf-8")
        self.assertIn('"classification": "PUBLIC_SYNTHETIC_ONLY"', text)
        self.assertIn('"publishable": false', text)

    def test_public_scope_is_synthetic(self) -> None:
        text = (ROOT / "config/scope/public_synthetic_scope.json").read_text(encoding="utf-8")
        self.assertIn('"classification": "PUBLIC_SYNTHETIC_ONLY"', text)

    def test_public_mapping_is_synthetic(self) -> None:
        text = (ROOT / "config/mapping/public_synthetic_mapping_v1.json").read_text(encoding="utf-8")
        self.assertIn('"classification": "PUBLIC_SYNTHETIC_ONLY"', text)


if __name__ == "__main__":
    unittest.main()
