"""Command-line entry point for scaffold and future pipeline commands."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from .activity.day5_pipeline import run_day5_upstream_rebuild
from .bootstrap import run_startup_checks
from .capability.pipeline import run_wp6_2_capability_detection
from .calculation.day7_pipeline import run_day7_calculation_and_lineage
from .cleaning.scope_filter import run_day3_scope_and_cleaning
from .errors import PipelineUserError
from .export.day8_export import run_day8_export
from .export.day8_run_all import run_day8_full_pipeline
from .factors.day6_pipeline import run_day6_factors_and_matching
from .io.excel_importer import inspect_excel_to_run
from .standardization.standardizer import run_day4_standardization
from .wp6_8_1.pipeline import run_end_to_end_pipeline
from .wp6_4.pipeline import run_wp6_4_validation
from .wp6_5.pipeline import WP65ValidationError, run_wp6_5_validation
from .wp6_6.pipeline import WP66AnalysisError, run_wp6_6_analysis


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="carbon-excel-pipeline")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("check", help="Validate project contracts and the scaffold.")
    inspect_parser = commands.add_parser(
        "inspect", help="Receive and inspect one .xlsx file in an isolated run."
    )
    inspect_parser.add_argument("--input", required=True, type=Path)
    inspect_parser.add_argument("--run-root", required=True, type=Path)
    inspect_parser.add_argument(
        "--alias-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "import" / "field_aliases.json",
    )
    inspect_parser.add_argument("--max-size-mb", type=int, default=50)
    capability_parser = commands.add_parser(
        "capability-detect",
        help="Detect record and dataset capabilities from WP6-1 structured artifacts.",
    )
    capability_parser.add_argument("--run-dir", required=True, type=Path)
    e2e_parser = commands.add_parser(
        "wp6-8-1-e2e",
        help="Run the WP6-8.1 capability-driven end-to-end pipeline on one Excel file.",
    )
    e2e_parser.add_argument("--input", required=True, type=Path)
    e2e_parser.add_argument("--run-root", required=True, type=Path)
    e2e_parser.add_argument("--private-id-baseline", type=Path)
    e2e_parser.add_argument("--private-mapping-workbook", type=Path)
    e2e_parser.add_argument("--standard-baseline", type=Path)
    e2e_parser.add_argument("--activity-baseline", type=Path)
    e2e_parser.add_argument("--third-party-baseline", type=Path)
    wp63_parser = commands.add_parser(
        "wp6-3-reproduce-2024",
        help="Consume WP6-1/2 artifacts and build the controlled 2024 reproduction.",
    )
    wp63_parser.add_argument("--upstream-run-dir", required=True, type=Path)
    wp63_parser.add_argument("--output-root", required=True, type=Path)
    wp63_parser.add_argument(
        "--policy-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "wp6" / "synthetic_direct_mass_policy.json",
    )
    wp63_parser.add_argument("--run-id")
    wp64_parser = commands.add_parser(
        "wp6-4-validate-2025",
        help="Validate the current 2025 run against WP5 and reconcile enterprise scope.",
    )
    wp64_parser.add_argument("--current-run-dir", required=True, type=Path)
    wp64_parser.add_argument("--frozen-result", required=True, type=Path)
    wp64_parser.add_argument("--frozen-lineage", required=True, type=Path)
    wp64_parser.add_argument("--frozen-activity", required=True, type=Path)
    wp64_parser.add_argument("--enterprise-workbook", required=True, type=Path)
    wp64_parser.add_argument("--wp6-3-run-dir", required=True, type=Path)
    wp64_parser.add_argument("--raw-input", required=True, type=Path)
    wp64_parser.add_argument("--output-root", required=True, type=Path)
    wp64_parser.add_argument(
        "--strict-contract",
        type=Path,
        default=PROJECT_ROOT / "config" / "contracts" / "wp5_strict_contract.json",
    )
    wp64_parser.add_argument("--run-id")
    wp65_parser = commands.add_parser(
        "wp6-5-validate-independent",
        help="Independently recalculate the formal 2024/2025 populations.",
    )
    wp65_parser.add_argument("--wp6-3-run-dir", required=True, type=Path)
    wp65_parser.add_argument("--wp6-4-run-dir", required=True, type=Path)
    wp65_parser.add_argument("--wp6-4-current-run-dir", required=True, type=Path)
    wp65_parser.add_argument("--output-root", required=True, type=Path)
    wp65_parser.add_argument("--documentation-root", type=Path)
    wp65_parser.add_argument("--run-id")
    wp66_parser = commands.add_parser(
        "wp6-6-analyze-factors",
        help="Build A/B/C/D historical factor scenarios from WP6-5 validation.",
    )
    wp66_parser.add_argument("--wp6-5-run-dir", required=True, type=Path)
    wp66_parser.add_argument("--output-root", required=True, type=Path)
    wp66_parser.add_argument("--documentation-root", type=Path)
    wp66_parser.add_argument("--run-id")
    scope_parser = commands.add_parser(
        "scope-clean",
        help="Confirm mappings, filter the configured scope and clean raw values.",
    )
    scope_parser.add_argument("--run-dir", required=True, type=Path)
    scope_parser.add_argument(
        "--scope-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "scope" / "public_synthetic_scope.json",
    )
    scope_parser.add_argument(
        "--unit-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "cleaning" / "unit_mappings.json",
    )
    standard_parser = commands.add_parser(
        "standardize",
        help="Convert Day 3 candidates to the frozen ordered 31-field structure.",
    )
    standard_parser.add_argument("--run-dir", required=True, type=Path)
    standard_parser.add_argument(
        "--profile-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "profiles" / "public_synthetic_profile.json",
    )
    standard_parser.add_argument(
        "--contract",
        type=Path,
        default=PROJECT_ROOT
        / "config"
        / "standardization"
        / "standard_31_contract.json",
    )
    standard_parser.add_argument(
        "--mapping-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "mapping" / "public_synthetic_mapping_v1.json",
    )
    standard_parser.add_argument("--private-id-baseline", type=Path)
    standard_parser.add_argument("--private-mapping-workbook", type=Path)
    upstream_parser = commands.add_parser(
        "upstream",
        help="Run Day 5 quality, activity, third-party input and the G1A gate.",
    )
    upstream_parser.add_argument("--run-dir", required=True, type=Path)
    upstream_parser.add_argument(
        "--profile-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "profiles" / "public_synthetic_profile.json",
    )
    upstream_parser.add_argument(
        "--standard-contract",
        type=Path,
        default=PROJECT_ROOT
        / "config"
        / "standardization"
        / "standard_31_contract.json",
    )
    upstream_parser.add_argument(
        "--activity-contract",
        type=Path,
        default=PROJECT_ROOT / "config" / "activity" / "activity_36_contract.json",
    )
    upstream_parser.add_argument(
        "--third-party-contract",
        type=Path,
        default=PROJECT_ROOT
        / "config"
        / "activity"
        / "third_party_20_contract.json",
    )
    upstream_parser.add_argument(
        "--quality-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "qc" / "day5_quality_rules.json",
    )
    upstream_parser.add_argument(
        "--interface-open-items",
        type=Path,
        default=PROJECT_ROOT
        / "config"
        / "open_items"
        / "wp3_interface_open_items.json",
    )
    upstream_parser.add_argument("--standard-baseline", type=Path)
    upstream_parser.add_argument("--activity-baseline", type=Path)
    upstream_parser.add_argument("--third-party-baseline", type=Path)
    factor_parser = commands.add_parser(
        "factor-match",
        help="Receive an eight-field factor file, adapt to D1 and exact-match by Record_ID.",
    )
    factor_parser.add_argument("--run-dir", required=True, type=Path)
    source_group = factor_parser.add_mutually_exclusive_group(required=True)
    source_group.add_argument("--factor-input", type=Path)
    source_group.add_argument("--historical-simulation", action="store_true")
    factor_parser.add_argument(
        "--profile-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "profiles" / "public_synthetic_profile.json",
    )
    factor_parser.add_argument(
        "--activity-contract",
        type=Path,
        default=PROJECT_ROOT / "config" / "activity" / "activity_36_contract.json",
    )
    factor_parser.add_argument(
        "--external-contract",
        type=Path,
        default=PROJECT_ROOT / "config" / "factors" / "external_8_contract.json",
    )
    factor_parser.add_argument(
        "--d1-contract",
        type=Path,
        default=PROJECT_ROOT / "config" / "factors" / "wp5_d1_45_contract.json",
    )
    factor_parser.add_argument(
        "--factor-config",
        type=Path,
        default=PROJECT_ROOT / "config" / "factors" / "public_synthetic_factor.json",
    )
    factor_parser.add_argument(
        "--d2-contract",
        type=Path,
        default=PROJECT_ROOT / "config" / "matching" / "wp5_d2_57_contract.json",
    )
    factor_parser.add_argument(
        "--route-contract",
        type=Path,
        default=PROJECT_ROOT / "config" / "matching" / "wp5_d3_route_36_contract.json",
    )
    calculation_parser = commands.add_parser(
        "calculate-lineage",
        help="Calculate D4/D5 results and build frozen plus extended lineage.",
    )
    calculation_parser.add_argument("--run-dir", required=True, type=Path)
    calculation_parser.add_argument(
        "--profile-config", type=Path,
        default=PROJECT_ROOT / "config" / "profiles" / "public_synthetic_profile.json",
    )
    calculation_parser.add_argument(
        "--calculation-rules", type=Path,
        default=PROJECT_ROOT / "config" / "calculation" / "day7_calculation_rules.json",
    )
    calculation_parser.add_argument(
        "--activity-contract", type=Path,
        default=PROJECT_ROOT / "config" / "activity" / "activity_36_contract.json",
    )
    calculation_parser.add_argument(
        "--d1-contract", type=Path,
        default=PROJECT_ROOT / "config" / "factors" / "wp5_d1_45_contract.json",
    )
    calculation_parser.add_argument(
        "--d2-contract", type=Path,
        default=PROJECT_ROOT / "config" / "matching" / "wp5_d2_57_contract.json",
    )
    calculation_parser.add_argument(
        "--d3-contract", type=Path,
        default=PROJECT_ROOT / "config" / "matching" / "wp5_d3_route_36_contract.json",
    )
    calculation_parser.add_argument(
        "--d4-contract", type=Path,
        default=PROJECT_ROOT / "config" / "calculation" / "wp5_d4_48_contract.json",
    )
    calculation_parser.add_argument(
        "--d5-contract", type=Path,
        default=PROJECT_ROOT / "config" / "calculation" / "wp5_d5_56_contract.json",
    )
    calculation_parser.add_argument(
        "--frozen-lineage-contract", type=Path,
        default=PROJECT_ROOT / "config" / "lineage" / "wp5_frozen_32_contract.json",
    )
    calculation_parser.add_argument(
        "--extended-lineage-contract", type=Path,
        default=PROJECT_ROOT / "config" / "lineage" / "demo_extended_lineage_contract.json",
    )
    export_parser = commands.add_parser(
        "export-workbook",
        help="Build and verify the Day 8 multi-sheet workbook for an existing passed run.",
    )
    export_parser.add_argument("--run-dir", required=True, type=Path)
    export_parser.add_argument("--output-dir", required=True, type=Path)
    export_parser.add_argument("--artifact-work-dir", required=True, type=Path)
    export_parser.add_argument("--node-executable", required=True, type=Path)
    export_parser.add_argument("--node-modules", required=True, type=Path)
    export_parser.add_argument(
        "--builder-script", type=Path,
        default=PROJECT_ROOT / "scripts" / "build_day8_workbook.mjs",
    )
    export_parser.add_argument("--wp5-open-items", required=True, type=Path)

    all_parser = commands.add_parser(
        "run-all",
        help="Run the complete private pipeline from raw Excel through the G2 workbook.",
    )
    all_parser.add_argument("--input", required=True, type=Path)
    all_parser.add_argument("--run-root", required=True, type=Path)
    all_parser.add_argument("--output-root", required=True, type=Path)
    all_parser.add_argument("--artifact-work-root", required=True, type=Path)
    all_parser.add_argument("--node-executable", required=True, type=Path)
    all_parser.add_argument("--node-modules", required=True, type=Path)
    all_parser.add_argument("--private-id-baseline", required=True, type=Path)
    all_parser.add_argument("--private-mapping-workbook", required=True, type=Path)
    all_parser.add_argument("--standard-baseline", required=True, type=Path)
    all_parser.add_argument("--activity-baseline", required=True, type=Path)
    all_parser.add_argument("--third-party-baseline", required=True, type=Path)
    all_parser.add_argument("--wp5-open-items", required=True, type=Path)
    all_parser.add_argument(
        "--builder-script", type=Path,
        default=PROJECT_ROOT / "scripts" / "build_day8_workbook.mjs",
    )
    all_parser.add_argument(
        "--alias-config", type=Path,
        default=PROJECT_ROOT / "config" / "import" / "field_aliases.json",
    )
    all_parser.add_argument(
        "--scope-config", type=Path,
        default=PROJECT_ROOT / "config" / "scope" / "public_synthetic_scope.json",
    )
    all_parser.add_argument(
        "--unit-config", type=Path,
        default=PROJECT_ROOT / "config" / "cleaning" / "unit_mappings.json",
    )
    all_parser.add_argument(
        "--profile-config", type=Path,
        default=PROJECT_ROOT / "config" / "profiles" / "public_synthetic_profile.json",
    )
    all_parser.add_argument(
        "--standard-contract", type=Path,
        default=PROJECT_ROOT / "config" / "standardization" / "standard_31_contract.json",
    )
    all_parser.add_argument(
        "--mapping-config", type=Path,
        default=PROJECT_ROOT / "config" / "mapping" / "public_synthetic_mapping_v1.json",
    )
    all_parser.add_argument(
        "--activity-contract", type=Path,
        default=PROJECT_ROOT / "config" / "activity" / "activity_36_contract.json",
    )
    all_parser.add_argument(
        "--third-party-contract", type=Path,
        default=PROJECT_ROOT / "config" / "activity" / "third_party_20_contract.json",
    )
    all_parser.add_argument(
        "--quality-config", type=Path,
        default=PROJECT_ROOT / "config" / "qc" / "day5_quality_rules.json",
    )
    all_parser.add_argument(
        "--interface-open-items", type=Path,
        default=PROJECT_ROOT / "config" / "open_items" / "wp3_interface_open_items.json",
    )
    all_parser.add_argument(
        "--external-factor-contract", type=Path,
        default=PROJECT_ROOT / "config" / "factors" / "external_8_contract.json",
    )
    all_parser.add_argument(
        "--d1-contract", type=Path,
        default=PROJECT_ROOT / "config" / "factors" / "wp5_d1_45_contract.json",
    )
    all_parser.add_argument(
        "--factor-config", type=Path,
        default=PROJECT_ROOT / "config" / "factors" / "public_synthetic_factor.json",
    )
    all_parser.add_argument(
        "--d2-contract", type=Path,
        default=PROJECT_ROOT / "config" / "matching" / "wp5_d2_57_contract.json",
    )
    all_parser.add_argument(
        "--d3-contract", type=Path,
        default=PROJECT_ROOT / "config" / "matching" / "wp5_d3_route_36_contract.json",
    )
    all_parser.add_argument(
        "--calculation-rules", type=Path,
        default=PROJECT_ROOT / "config" / "calculation" / "day7_calculation_rules.json",
    )
    all_parser.add_argument(
        "--d4-contract", type=Path,
        default=PROJECT_ROOT / "config" / "calculation" / "wp5_d4_48_contract.json",
    )
    all_parser.add_argument(
        "--d5-contract", type=Path,
        default=PROJECT_ROOT / "config" / "calculation" / "wp5_d5_56_contract.json",
    )
    all_parser.add_argument(
        "--frozen-lineage-contract", type=Path,
        default=PROJECT_ROOT / "config" / "lineage" / "wp5_frozen_32_contract.json",
    )
    all_parser.add_argument(
        "--extended-lineage-contract", type=Path,
        default=PROJECT_ROOT / "config" / "lineage" / "demo_extended_lineage_contract.json",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "check":
        result = run_startup_checks()
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 1
    if args.command == "inspect":
        try:
            result = inspect_excel_to_run(
                args.input,
                run_root=args.run_root,
                alias_config_path=args.alias_config,
                max_size_bytes=args.max_size_mb * 1024 * 1024,
            )
        except PipelineUserError as error:
            print(json.dumps({"status": "BLOCKED", "error": error.to_dict()}, ensure_ascii=False, indent=2))
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "capability-detect":
        try:
            result = run_wp6_2_capability_detection(args.run_dir)
        except PipelineUserError as error:
            print(
                json.dumps(
                    {"status": "BLOCKED", "error": error.to_dict()},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    if args.command == "wp6-8-1-e2e":
        private_paths = None
        required = (
            args.private_id_baseline,
            args.private_mapping_workbook,
            args.standard_baseline,
            args.activity_baseline,
            args.third_party_baseline,
        )
        if all(required):
            private_paths = {
                "private_id_baseline": args.private_id_baseline,
                "private_mapping_workbook": args.private_mapping_workbook,
                "standard_baseline": args.standard_baseline,
                "activity_baseline": args.activity_baseline,
                "third_party_baseline": args.third_party_baseline,
            }
        try:
            result = run_end_to_end_pipeline(
                args.input,
                run_root=args.run_root,
                private_paths=private_paths,
            )
        except PipelineUserError as error:
            print(json.dumps({"status": "BLOCKED", "error": error.to_dict()}, ensure_ascii=False, indent=2))
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
        return 0 if result["status"] not in {"BLOCKED", "FAILED"} else 2
    if args.command == "wp6-3-reproduce-2024":
        try:
            result = run_wp6_3_historical_reproduction(
                args.upstream_run_dir,
                output_root=args.output_root,
                policy_config_path=args.policy_config,
                run_id=args.run_id,
            )
        except PipelineUserError as error:
            print(
                json.dumps(
                    {"status": "BLOCKED", "error": error.to_dict()},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] in {"PASS", "PASS_WITH_WARNING"} else 2
    if args.command == "wp6-4-validate-2025":
        try:
            result = run_wp6_4_validation(
                args.current_run_dir,
                frozen_result_path=args.frozen_result,
                frozen_lineage_path=args.frozen_lineage,
                frozen_activity_path=args.frozen_activity,
                enterprise_workbook_path=args.enterprise_workbook,
                wp6_3_run_dir=args.wp6_3_run_dir,
                raw_input_path=args.raw_input,
                output_root=args.output_root,
                strict_contract_path=args.strict_contract,
                run_id=args.run_id,
            )
        except PipelineUserError as error:
            print(
                json.dumps(
                    {"status": "BLOCKED", "error": error.to_dict()},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "wp6-5-validate-independent":
        try:
            result = run_wp6_5_validation(
                wp6_3_run_dir=args.wp6_3_run_dir,
                wp6_4_run_dir=args.wp6_4_run_dir,
                wp6_4_current_run_dir=args.wp6_4_current_run_dir,
                output_root=args.output_root,
                documentation_root=args.documentation_root,
                run_id=args.run_id,
            )
        except WP65ValidationError as error:
            print(
                json.dumps(
                    {"stage": "WP6-5", "status": "BLOCKED", "error": str(error)},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "wp6-6-analyze-factors":
        try:
            result = run_wp6_6_analysis(
                wp6_5_run_dir=args.wp6_5_run_dir,
                output_root=args.output_root,
                documentation_root=args.documentation_root,
                run_id=args.run_id,
            )
        except WP66AnalysisError as error:
            print(
                json.dumps(
                    {"stage": "WP6-6", "status": "BLOCKED", "error": str(error)},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "scope-clean":
        try:
            result = run_day3_scope_and_cleaning(
                args.run_dir,
                scope_config_path=args.scope_config,
                unit_config_path=args.unit_config,
            )
        except PipelineUserError as error:
            print(
                json.dumps(
                    {"status": "BLOCKED", "error": error.to_dict()},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "standardize":
        try:
            result = run_day4_standardization(
                args.run_dir,
                profile_config_path=args.profile_config,
                contract_path=args.contract,
                mapping_config_path=args.mapping_config,
                private_id_baseline_path=args.private_id_baseline,
                private_mapping_workbook_path=args.private_mapping_workbook,
            )
        except PipelineUserError as error:
            print(
                json.dumps(
                    {"status": "BLOCKED", "error": error.to_dict()},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "upstream":
        try:
            result = run_day5_upstream_rebuild(
                args.run_dir,
                profile_config_path=args.profile_config,
                standard_contract_path=args.standard_contract,
                activity_contract_path=args.activity_contract,
                third_party_contract_path=args.third_party_contract,
                quality_config_path=args.quality_config,
                interface_open_items_path=args.interface_open_items,
                standard_baseline_path=args.standard_baseline,
                activity_baseline_path=args.activity_baseline,
                third_party_baseline_path=args.third_party_baseline,
            )
        except PipelineUserError as error:
            print(
                json.dumps(
                    {"status": "BLOCKED", "error": error.to_dict()},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "factor-match":
        try:
            result = run_day6_factors_and_matching(
                args.run_dir,
                profile_config_path=args.profile_config,
                activity_contract_path=args.activity_contract,
                external_contract_path=args.external_contract,
                d1_contract_path=args.d1_contract,
                factor_config_path=args.factor_config,
                d2_contract_path=args.d2_contract,
                route_contract_path=args.route_contract,
                factor_input_path=args.factor_input,
                historical_simulation=args.historical_simulation,
            )
        except PipelineUserError as error:
            print(
                json.dumps(
                    {"status": "BLOCKED", "error": error.to_dict()},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "calculate-lineage":
        try:
            result = run_day7_calculation_and_lineage(
                args.run_dir,
                profile_config_path=args.profile_config,
                calculation_rules_path=args.calculation_rules,
                activity_contract_path=args.activity_contract,
                d1_contract_path=args.d1_contract,
                d2_contract_path=args.d2_contract,
                d3_contract_path=args.d3_contract,
                d4_contract_path=args.d4_contract,
                d5_contract_path=args.d5_contract,
                frozen_lineage_contract_path=args.frozen_lineage_contract,
                extended_lineage_contract_path=args.extended_lineage_contract,
            )
        except PipelineUserError as error:
            print(json.dumps({"status": "BLOCKED", "error": error.to_dict()}, ensure_ascii=False, indent=2))
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "export-workbook":
        try:
            result = run_day8_export(
                args.run_dir,
                output_dir=args.output_dir,
                artifact_work_dir=args.artifact_work_dir,
                node_executable=args.node_executable,
                node_modules_path=args.node_modules,
                builder_script_path=args.builder_script,
                wp5_open_items_path=args.wp5_open_items,
            )
        except PipelineUserError as error:
            print(json.dumps({"status": "BLOCKED", "error": error.to_dict()}, ensure_ascii=False, indent=2))
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    if args.command == "run-all":
        try:
            result = run_day8_full_pipeline(
                args.input,
                run_root=args.run_root,
                output_root=args.output_root,
                artifact_work_root=args.artifact_work_root,
                node_executable=args.node_executable,
                node_modules_path=args.node_modules,
                builder_script_path=args.builder_script,
                alias_config_path=args.alias_config,
                scope_config_path=args.scope_config,
                unit_config_path=args.unit_config,
                profile_config_path=args.profile_config,
                standard_contract_path=args.standard_contract,
                mapping_config_path=args.mapping_config,
                private_id_baseline_path=args.private_id_baseline,
                private_mapping_workbook_path=args.private_mapping_workbook,
                activity_contract_path=args.activity_contract,
                third_party_contract_path=args.third_party_contract,
                quality_config_path=args.quality_config,
                interface_open_items_path=args.interface_open_items,
                standard_baseline_path=args.standard_baseline,
                activity_baseline_path=args.activity_baseline,
                third_party_baseline_path=args.third_party_baseline,
                external_factor_contract_path=args.external_factor_contract,
                d1_contract_path=args.d1_contract,
                factor_config_path=args.factor_config,
                d2_contract_path=args.d2_contract,
                d3_contract_path=args.d3_contract,
                calculation_rules_path=args.calculation_rules,
                d4_contract_path=args.d4_contract,
                d5_contract_path=args.d5_contract,
                frozen_lineage_contract_path=args.frozen_lineage_contract,
                extended_lineage_contract_path=args.extended_lineage_contract,
                wp5_open_items_path=args.wp5_open_items,
            )
        except PipelineUserError as error:
            print(json.dumps({"status": "BLOCKED", "error": error.to_dict()}, ensure_ascii=False, indent=2))
            return 2
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 2
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
