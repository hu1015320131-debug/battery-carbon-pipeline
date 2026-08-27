"""Validate public contracts without reading business data."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.export.day8_validation import EXPECTED_SHEET_NAMES

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _load_json(relative_path: str) -> dict[str, Any]:
    path = PROJECT_ROOT / relative_path
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def run_startup_checks() -> dict[str, Any]:
    errors: list[str] = []
    contract = _load_json("config/contracts/wp5_strict_contract.json")
    registry = _load_json("config/profiles/profile_registry.json")

    expected_counts = {
        "D1": 45,
        "D2": 57,
        "D3": 36,
        "D4": 48,
        "D5": 56,
        "FROZEN_LINEAGE": 32,
    }
    if contract.get("stage_field_counts") != expected_counts:
        errors.append("Stage field counts must remain 45/57/36/48/56/32.")

    if contract.get("synthetic_factor", {}).get("value") != "1.250000":
        errors.append("The synthetic demonstration factor must remain 1.250000.")

    required_units = {
        "PCS",
        "g/PCS",
        "g/year",
        "kg/year",
        "kgCO2e/kg",
        "kgCO2e/year",
    }
    if set(contract.get("machine_units", [])) != required_units:
        errors.append("The machine-unit set is incomplete or changed.")

    profile_ids = registry.get("profiles", [])
    if profile_ids != ["public_synthetic_profile"]:
        errors.append("Profile registry must contain only public_synthetic_profile.")

    public = _load_json("config/profiles/public_synthetic_profile.json")
    if public.get("classification") != "PUBLIC_SYNTHETIC_ONLY":
        errors.append("The public profile must remain PUBLIC_SYNTHETIC_ONLY.")
    if public.get("production_eligible") is True:
        errors.append("The public profile cannot be production eligible.")

    public_scope = _load_json("config/scope/public_synthetic_scope.json")
    if public_scope.get("classification") != "PUBLIC_SYNTHETIC_ONLY":
        errors.append("Public scope configuration must remain synthetic-only.")

    public_mapping = _load_json("config/mapping/public_synthetic_mapping_v1.json")
    if public_mapping.get("classification") != "PUBLIC_SYNTHETIC_ONLY":
        errors.append("Public mapping configuration must remain synthetic-only.")

    standard_contract = _load_json("config/standardization/standard_31_contract.json")
    if standard_contract.get("field_count") != 31:
        errors.append("The Day 4 standard contract must contain exactly 31 fields.")

    activity_contract = _load_json("config/activity/activity_36_contract.json")
    third_party_contract = _load_json("config/activity/third_party_20_contract.json")
    if activity_contract.get("field_count") != 36:
        errors.append("The Day 5 activity contract must contain exactly 36 fields.")
    if third_party_contract.get("field_count") != 20:
        errors.append("The Day 5 third-party contract must contain exactly 20 fields.")

    day6_contracts = {
        "config/factors/external_8_contract.json": 8,
        "config/factors/wp5_d1_45_contract.json": 45,
        "config/matching/wp5_d2_57_contract.json": 57,
        "config/matching/wp5_d3_route_36_contract.json": 36,
    }
    for path, expected in day6_contracts.items():
        day6_contract = _load_json(path)
        if day6_contract.get("field_count") != expected:
            errors.append(f"The Day 6 contract {path} must contain {expected} fields.")

    factor_config = _load_json("config/factors/public_synthetic_factor.json")
    if (factor_config.get("ef_value"), factor_config.get("ef_unit")) != ("1.250000", "kgCO2e/kg"):
        errors.append("The synthetic factor fixture is not 1.250000 kgCO2e/kg.")
    if factor_config.get("production_eligible") != "FALSE":
        errors.append("The synthetic factor cannot be production eligible.")

    if len(EXPECTED_SHEET_NAMES) != 18 or len(set(EXPECTED_SHEET_NAMES)) != 18:
        errors.append("The Day 8 workbook must contain 18 unique ordered sheets.")

    required_dirs = [
        "app",
        "config",
        "scripts",
        "tests/unit",
        "tests/integration",
        "tests/public",
        "src/carbon_excel_pipeline/io",
        "src/carbon_excel_pipeline/cleaning",
        "src/carbon_excel_pipeline/qc",
        "src/carbon_excel_pipeline/activity",
        "src/carbon_excel_pipeline/factors",
        "src/carbon_excel_pipeline/matching",
        "src/carbon_excel_pipeline/calculation",
        "src/carbon_excel_pipeline/export",
        "examples/public",
    ]
    missing_dirs = [item for item in required_dirs if not (PROJECT_ROOT / item).is_dir()]
    if missing_dirs:
        errors.append(f"Missing scaffold directories: {missing_dirs}")

    return {
        "status": "PASS" if not errors else "FAIL",
        "gate_candidate": "G0_READY_TO_BUILD" if not errors else "G0_BLOCKED",
        "business_pipeline_executed": False,
        "real_data_read": False,
        "profiles": profile_ids,
        "errors": errors,
    }
