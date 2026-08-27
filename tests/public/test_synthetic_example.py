from __future__ import annotations

import hashlib
from pathlib import Path

from carbon_excel_pipeline.wp6_8_1.router import (
    ACTIVITY_DIRECT,
    FACTOR_EMBEDDED,
    decide_processing_routes,
)


ROOT = Path(__file__).resolve().parents[2]
EXAMPLE = ROOT / "examples" / "public" / "synthetic_cells.xlsx"


def test_synthetic_example_workbook_exists_and_is_small() -> None:
    assert EXAMPLE.is_file()
    assert EXAMPLE.stat().st_size < 50_000


def test_synthetic_example_sha_is_the_only_gated_boundary() -> None:
    digest = hashlib.sha256(EXAMPLE.read_bytes()).hexdigest().upper()
    decision = decide_processing_routes(
        capability={
            "direct_reported_mass_count": 2,
            "pcs_weight_derived_count": 0,
            "factor_ready_count": 2,
        },
        input_sha256=digest,
    )
    assert decision["Activity_Route"] == ACTIVITY_DIRECT
    assert decision["Factor_Route"] == FACTOR_EMBEDDED
    assert decision["Boundary_Policy"] == "SYNTHETIC_DIRECT_MASS_BOUNDARY_V1"
    assert decision["status"] == "ROUTED"
    assert decision["Production_Eligible"] is False
