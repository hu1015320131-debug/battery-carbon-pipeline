"""One-command Day 8 orchestration from the raw Excel workbook to G2 output."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from carbon_excel_pipeline.activity.day5_pipeline import run_day5_upstream_rebuild
from carbon_excel_pipeline.calculation.day7_pipeline import run_day7_calculation_and_lineage
from carbon_excel_pipeline.cleaning.scope_filter import run_day3_scope_and_cleaning
from carbon_excel_pipeline.export.day8_export import run_day8_export
from carbon_excel_pipeline.factors.day6_pipeline import run_day6_factors_and_matching
from carbon_excel_pipeline.io.excel_importer import inspect_excel_to_run
from carbon_excel_pipeline.standardization.standardizer import run_day4_standardization


def run_day8_full_pipeline(
    input_path: Path,
    *,
    run_root: Path,
    output_root: Path,
    artifact_work_root: Path,
    node_executable: Path,
    node_modules_path: Path,
    builder_script_path: Path,
    alias_config_path: Path,
    scope_config_path: Path,
    unit_config_path: Path,
    profile_config_path: Path,
    standard_contract_path: Path,
    mapping_config_path: Path,
    private_id_baseline_path: Path,
    private_mapping_workbook_path: Path,
    activity_contract_path: Path,
    third_party_contract_path: Path,
    quality_config_path: Path,
    interface_open_items_path: Path,
    standard_baseline_path: Path,
    activity_baseline_path: Path,
    third_party_baseline_path: Path,
    external_factor_contract_path: Path,
    d1_contract_path: Path,
    factor_config_path: Path,
    d2_contract_path: Path,
    d3_contract_path: Path,
    calculation_rules_path: Path,
    d4_contract_path: Path,
    d5_contract_path: Path,
    frozen_lineage_contract_path: Path,
    extended_lineage_contract_path: Path,
    wp5_open_items_path: Path,
) -> dict[str, Any]:
    day2 = inspect_excel_to_run(
        input_path,
        run_root=run_root,
        alias_config_path=alias_config_path,
        max_size_bytes=50 * 1024 * 1024,
    )
    run_dir = Path(str(day2["run_directory"])).resolve()
    day3 = run_day3_scope_and_cleaning(
        run_dir,
        scope_config_path=scope_config_path,
        unit_config_path=unit_config_path,
    )
    day4 = run_day4_standardization(
        run_dir,
        profile_config_path=profile_config_path,
        contract_path=standard_contract_path,
        mapping_config_path=mapping_config_path,
        private_id_baseline_path=private_id_baseline_path,
        private_mapping_workbook_path=private_mapping_workbook_path,
    )
    day5 = run_day5_upstream_rebuild(
        run_dir,
        profile_config_path=profile_config_path,
        standard_contract_path=standard_contract_path,
        activity_contract_path=activity_contract_path,
        third_party_contract_path=third_party_contract_path,
        quality_config_path=quality_config_path,
        interface_open_items_path=interface_open_items_path,
        standard_baseline_path=standard_baseline_path,
        activity_baseline_path=activity_baseline_path,
        third_party_baseline_path=third_party_baseline_path,
    )
    day6 = run_day6_factors_and_matching(
        run_dir,
        profile_config_path=profile_config_path,
        activity_contract_path=activity_contract_path,
        external_contract_path=external_factor_contract_path,
        d1_contract_path=d1_contract_path,
        factor_config_path=factor_config_path,
        d2_contract_path=d2_contract_path,
        route_contract_path=d3_contract_path,
        historical_simulation=True,
    )
    day7 = run_day7_calculation_and_lineage(
        run_dir,
        profile_config_path=profile_config_path,
        calculation_rules_path=calculation_rules_path,
        activity_contract_path=activity_contract_path,
        d1_contract_path=d1_contract_path,
        d2_contract_path=d2_contract_path,
        d3_contract_path=d3_contract_path,
        d4_contract_path=d4_contract_path,
        d5_contract_path=d5_contract_path,
        frozen_lineage_contract_path=frozen_lineage_contract_path,
        extended_lineage_contract_path=extended_lineage_contract_path,
    )
    day8 = run_day8_export(
        run_dir,
        output_dir=output_root / run_dir.name,
        artifact_work_dir=artifact_work_root / run_dir.name,
        node_executable=node_executable,
        node_modules_path=node_modules_path,
        builder_script_path=builder_script_path,
        wp5_open_items_path=wp5_open_items_path,
    )
    return {
        **day8,
        "command": "run-all",
        "stage_statuses": {
            "day2": day2["status"],
            "day3": day3["status"],
            "day4": day4["status"],
            "day5": day5["status"],
            "day6": day6["status"],
            "day7": day7["status"],
            "day8": day8["status"],
        },
    }
