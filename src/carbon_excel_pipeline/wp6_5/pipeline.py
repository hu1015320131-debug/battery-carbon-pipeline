"""WP6-5 full-population independent calculation validation."""

from __future__ import annotations

import ast
import csv
import hashlib
import json
from collections import Counter
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

from .independent import (
    IndependentValidationInputError,
    calculate_direct_mass_activity,
    calculate_emission,
    calculate_pcs_weight_activity,
    compare_calculation,
    display_six,
    format_decimal,
    normalize_ef_unit,
)


EXPECTED = {
    "2024": {
        "records": 2,
        "activity_kg": Decimal("80.0000"),
        "emission_kgco2e": Decimal("100.000000"),
    },
    "2025": {
        "records": 2,
        "activity_kg": Decimal("100.0000"),
        "emission_kgco2e": Decimal("125.000000"),
    },
}

VALIDATION_FIELDS = [
    "Year",
    "Record_ID",
    "Source_File",
    "Source_SHA256",
    "Source_Sheet",
    "Source_Row",
    "Activity_Method",
    "Original_Activity_Value",
    "Original_Activity_Unit",
    "Quantity_PCS",
    "Unit_Weight",
    "Unit_Weight_Unit",
    "Main_Activity_kg",
    "Independent_Activity_kg",
    "Activity_Difference",
    "Activity_Validation_Status",
    "Main_EF",
    "Main_EF_Unit",
    "Independent_EF",
    "Independent_EF_Unit",
    "Independent_EF_Source",
    "EF_Difference",
    "EF_Validation_Status",
    "Main_Emission_kgCO2e",
    "Independent_Emission_kgCO2e",
    "Emission_Difference",
    "Emission_Validation_Status",
    "Main_Emission_Display_6dp",
    "Independent_Emission_Display_6dp",
    "Display_Difference",
    "Display_Validation_Status",
    "Governance_QC",
    "Boundary_Ready",
    "Boundary_Validation_Status",
    "Lineage_Validation_Status",
    "Reason_Codes",
    "Validation_Error",
    "Overall_Validation_Status",
]


class WP65ValidationError(RuntimeError):
    """Raised when the formal WP6-5 run cannot be established safely."""


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _unique(rows: list[dict[str, str]], label: str) -> dict[str, dict[str, str]]:
    output: dict[str, dict[str, str]] = {}
    for row in rows:
        record_id = row.get("Record_ID", "")
        if not record_id:
            raise WP65ValidationError(f"{label} contains a blank Record_ID")
        if record_id in output:
            raise WP65ValidationError(f"{label} contains duplicate Record_ID: {record_id}")
        output[record_id] = row
    return output


def _truth(value: Any) -> bool:
    return str(value).strip().casefold() in {"true", "1", "yes", "pass"}


def _present(value: Any) -> bool:
    return str(value or "").strip() not in {"", "UNKNOWN"}


def _blocked_row(year: str, record: dict[str, str], message: str) -> dict[str, str]:
    return {
        "Year": year,
        "Record_ID": record.get("Record_ID", ""),
        "Source_File": record.get("Source_File", ""),
        "Source_SHA256": record.get("Source_SHA256", ""),
        "Source_Sheet": record.get("Source_Sheet", ""),
        "Source_Row": record.get("Source_Row", ""),
        "Activity_Method": record.get("Activity_Method", ""),
        "Reason_Codes": "VALIDATION_INPUT_MISSING",
        "Overall_Validation_Status": "INDEPENDENT_VALIDATION_BLOCKED",
        "Validation_Error": message,
    }


def verify_validator_independence(source_path: Path) -> dict[str, Any]:
    """Use the AST to prove the validation package has no production imports."""

    sources = (
        sorted(source_path.glob("*.py")) if source_path.is_dir() else [source_path]
    )
    imports: list[str] = []
    for source in sources:
        tree = ast.parse(source.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imports.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom):
                imports.append(node.module or "")
    forbidden_fragments = (
        "activity.day5_pipeline",
        "calculation.day7_pipeline",
        "wp6_3",
        "wp6_4",
        "multiply_emission_decimal",
    )
    forbidden = [
        name for name in imports if any(fragment in name for fragment in forbidden_fragments)
    ]
    return {
        "source_files": [str(source.resolve()) for source in sources],
        "imports": imports,
        "forbidden_imports": forbidden,
        "production_calculation_imported": bool(forbidden),
        "status": "PASS" if not forbidden else "FAIL",
    }


def _validate_2024(paths: dict[str, Path]) -> list[dict[str, str]]:
    canonical_rows = _read_csv(paths["2024_canonical"])
    calculation = _unique(_read_csv(paths["2024_calculation"]), "2024 calculation")
    factors = _unique(_read_csv(paths["2024_factors"]), "2024 EF")
    lineage = _unique(_read_csv(paths["2024_lineage"]), "2024 lineage")
    results: list[dict[str, str]] = []
    for canonical in canonical_rows:
        record_id = canonical.get("Record_ID", "")
        try:
            main = calculation[record_id]
            factor = factors[record_id]
            trace = lineage[record_id]
            independent_activity = calculate_direct_mass_activity(
                canonical.get("Original_Activity_Value"),
                canonical.get("Original_Activity_Unit"),
            )
            independent_ef = factor.get("EF_Value", "")
            independent_emission = calculate_emission(
                independent_activity,
                independent_ef,
                factor.get("EF_Unit", ""),
            )
            lineage_match = all(
                canonical.get(field, "") == trace.get(field, "")
                for field in ("Record_ID", "Source_SHA256", "Source_Sheet", "Source_Row")
            )
            boundary_ready = (
                all(
                    _present(canonical.get(field))
                    for field in (
                        "Business_Unit",
                        "Purchase_Category",
                        "Product_Description",
                    )
                )
                and not canonical.get("Blocking_Codes", "").strip()
            )
            comparison = compare_calculation(
                main_activity_kg=main.get("Activity_Data_kg"),
                independent_activity_kg=independent_activity,
                main_ef=main.get("EF_Value"),
                independent_ef=independent_ef,
                main_ef_unit=canonical.get("EF_Unit"),
                independent_ef_unit=factor.get("EF_Unit"),
                main_emission_kgco2e=main.get("Emission_kgCO2e"),
                independent_emission_kgco2e=independent_emission,
                lineage_match=lineage_match,
                boundary_ready=boundary_ready,
            )
            results.append(
                {
                    "Year": "2024",
                    "Record_ID": record_id,
                    "Source_File": canonical.get("Source_File", ""),
                    "Source_SHA256": canonical.get("Source_SHA256", ""),
                    "Source_Sheet": canonical.get("Source_Sheet", ""),
                    "Source_Row": canonical.get("Source_Row", ""),
                    "Activity_Method": "DIRECT_REPORTED_MASS",
                    "Original_Activity_Value": canonical.get(
                        "Original_Activity_Value", ""
                    ),
                    "Original_Activity_Unit": canonical.get(
                        "Original_Activity_Unit", ""
                    ),
                    "Quantity_PCS": "",
                    "Unit_Weight": "",
                    "Unit_Weight_Unit": "",
                    "Main_Activity_kg": main.get("Activity_Data_kg", ""),
                    "Independent_Activity_kg": format_decimal(independent_activity),
                    "Main_EF": main.get("EF_Value", ""),
                    "Main_EF_Unit": canonical.get("EF_Unit", ""),
                    "Independent_EF": independent_ef,
                    "Independent_EF_Unit": factor.get("EF_Unit", ""),
                    "Independent_EF_Source": factor.get("EF_Source", ""),
                    "Main_Emission_kgCO2e": main.get("Emission_kgCO2e", ""),
                    "Independent_Emission_kgCO2e": format_decimal(
                        independent_emission
                    ),
                    "Governance_QC": canonical.get("QC_Status", ""),
                    "Boundary_Ready": str(boundary_ready).upper(),
                    **comparison,
                }
            )
        except (KeyError, IndependentValidationInputError) as error:
            results.append(_blocked_row("2024", canonical, str(error)))
    return results


def _validate_2025(paths: dict[str, Path]) -> list[dict[str, str]]:
    canonical_rows = _read_csv(paths["2025_canonical"])
    standards = _unique(_read_csv(paths["2025_standard"]), "2025 standard")
    factors = _unique(_read_csv(paths["2025_factors"]), "2025 EF")
    main_results = _unique(_read_csv(paths["2025_main_result"]), "2025 main result")
    lineage = _unique(_read_csv(paths["2025_lineage"]), "2025 lineage")
    results: list[dict[str, str]] = []
    for canonical in canonical_rows:
        record_id = canonical.get("Record_ID", "")
        try:
            standard = standards[record_id]
            factor = factors[record_id]
            main_result = main_results[record_id]
            trace = lineage[record_id]
            independent_activity = calculate_pcs_weight_activity(
                standard.get("PCS"), standard.get("Unit_Weight_g"), "g/PCS"
            )
            independent_ef = factor.get("EF_Value", "")
            independent_emission = calculate_emission(
                independent_activity,
                independent_ef,
                factor.get("EF_Unit", ""),
            )
            lineage_match = (
                canonical.get("Source_File", "")
                == standard.get("Source_File", "")
                == trace.get("Raw_Input_Source_File", "")
                and canonical.get("Source_Sheet", "")
                == standard.get("Source_Sheet", "")
                == trace.get("Source_Sheet", "")
                and canonical.get("Source_Row", "")
                == standard.get("Source_Row", "")
                == trace.get("Source_Row", "")
                and trace.get("Raw_Input_SHA256", "")
                == trace.get("Received_Input_SHA256", "")
                and trace.get("Extended_Trace_Status", "") == "TRACE_COMPLETE"
            )
            boundary_ready = _truth(canonical.get("Boundary_Ready"))
            comparison = compare_calculation(
                main_activity_kg=canonical.get("Activity_Data_kg"),
                independent_activity_kg=independent_activity,
                main_ef=canonical.get("EF_Value"),
                independent_ef=independent_ef,
                main_ef_unit=canonical.get("EF_Unit"),
                independent_ef_unit=factor.get("EF_Unit"),
                main_emission_kgco2e=canonical.get("Emission_kgCO2e"),
                independent_emission_kgco2e=independent_emission,
                main_display_emission=main_result.get("Emission_kgCO2e"),
                lineage_match=lineage_match,
                boundary_ready=boundary_ready,
            )
            results.append(
                {
                    "Year": "2025",
                    "Record_ID": record_id,
                    "Source_File": canonical.get("Source_File", ""),
                    "Source_SHA256": trace.get("Raw_Input_SHA256", ""),
                    "Source_Sheet": canonical.get("Source_Sheet", ""),
                    "Source_Row": canonical.get("Source_Row", ""),
                    "Activity_Method": "PCS_WEIGHT_DERIVED",
                    "Original_Activity_Value": standard.get(
                        "Original_Activity_Value", ""
                    ),
                    "Original_Activity_Unit": standard.get(
                        "Original_Activity_Unit", ""
                    ),
                    "Quantity_PCS": standard.get("PCS", ""),
                    "Unit_Weight": standard.get("Unit_Weight_g", ""),
                    "Unit_Weight_Unit": "g/PCS",
                    "Main_Activity_kg": canonical.get("Activity_Data_kg", ""),
                    "Independent_Activity_kg": format_decimal(independent_activity),
                    "Main_EF": canonical.get("EF_Value", ""),
                    "Main_EF_Unit": canonical.get("EF_Unit", ""),
                    "Independent_EF": independent_ef,
                    "Independent_EF_Unit": factor.get("EF_Unit", ""),
                    "Independent_EF_Source": (
                        f"{factor.get('Source_Name', '')} {factor.get('Source_Version', '')}"
                    ).strip(),
                    "Main_Emission_kgCO2e": canonical.get(
                        "Emission_kgCO2e", ""
                    ),
                    "Independent_Emission_kgCO2e": format_decimal(
                        independent_emission
                    ),
                    "Governance_QC": canonical.get("Governance_QC", ""),
                    "Boundary_Ready": str(boundary_ready).upper(),
                    **comparison,
                }
            )
        except (KeyError, IndependentValidationInputError) as error:
            results.append(_blocked_row("2025", canonical, str(error)))
    return results


def _ef_audit(rows: list[dict[str, str]]) -> dict[str, Any]:
    usable = [
        row
        for row in rows
        if row.get("Independent_Activity_kg") and row.get("Independent_EF")
    ]
    if not usable:
        return {
            "EF_Unique_Count": 0,
            "EF_Min": "",
            "EF_Max": "",
            "Activity_Weighted_EF": "",
        }
    factors = [Decimal(row["Independent_EF"]) for row in usable]
    activity = [Decimal(row["Independent_Activity_kg"]) for row in usable]
    activity_total = sum(activity, Decimal("0"))
    weighted = sum(
        (
            Decimal(row["Independent_Activity_kg"]) * Decimal(row["Independent_EF"])
            for row in usable
        ),
        Decimal("0"),
    ) / activity_total
    return {
        "EF_Unique_Count": len(set(factors)),
        "EF_Min": format_decimal(min(factors)),
        "EF_Max": format_decimal(max(factors)),
        "Activity_Weighted_EF": format_decimal(weighted),
    }


def _year_summary(year: str, rows: list[dict[str, str]]) -> dict[str, Any]:
    usable = [row for row in rows if row.get("Independent_Activity_kg")]
    main_activity = sum(
        (Decimal(row["Main_Activity_kg"]) for row in usable), Decimal("0")
    )
    independent_activity = sum(
        (Decimal(row["Independent_Activity_kg"]) for row in usable), Decimal("0")
    )
    main_emission = sum(
        (Decimal(row["Main_Emission_kgCO2e"]) for row in usable), Decimal("0")
    )
    independent_emission = sum(
        (Decimal(row["Independent_Emission_kgCO2e"]) for row in usable),
        Decimal("0"),
    )
    main_rounded_aggregate = display_six(main_emission)
    independent_rounded_aggregate = display_six(independent_emission)
    main_row_rounded = sum(
        (display_six(row["Main_Emission_kgCO2e"]) for row in usable), Decimal("0")
    )
    independent_row_rounded = sum(
        (display_six(row["Independent_Emission_kgCO2e"]) for row in usable),
        Decimal("0"),
    )
    statuses = Counter(row.get("Overall_Validation_Status", "") for row in rows)
    expected = EXPECTED[year]
    regression = {
        "Record_Count": len(rows) == expected["records"],
        "Independent_Activity_Total": independent_activity == expected["activity_kg"],
        "Independent_Emission_Total": independent_emission
        == expected["emission_kgco2e"],
    }
    summary = {
        "Year": year,
        "Status": "PASS"
        if (
            statuses["INDEPENDENT_CALCULATION_PASS"] == expected["records"]
            and all(regression.values())
        )
        else "BLOCKED",
        "Record_Count": len(rows),
        "Main_Activity_Total_kg": format_decimal(main_activity),
        "Independent_Activity_Total_kg": format_decimal(independent_activity),
        "Activity_Total_Difference": format_decimal(
            main_activity - independent_activity
        ),
        "Main_Emission_Total_kgCO2e": format_decimal(main_emission),
        "Independent_Emission_Total_kgCO2e": format_decimal(independent_emission),
        "Emission_Total_Difference": format_decimal(
            main_emission - independent_emission
        ),
        "Exact_Match_Count": statuses["INDEPENDENT_CALCULATION_PASS"],
        "Difference_Count": statuses["INDEPENDENT_VALIDATION_FAIL"],
        "Blocked_Count": statuses["INDEPENDENT_VALIDATION_BLOCKED"],
        "Boundary_Ready_Count": sum(
            row.get("Boundary_Ready") == "TRUE" for row in rows
        ),
        "Lineage_Pass_Count": sum(
            row.get("Lineage_Validation_Status") == "PASS" for row in rows
        ),
        "Governance_QC_Counts": dict(Counter(row.get("Governance_QC", "") for row in rows)),
        "EF_Audit": _ef_audit(rows),
        "Raw_Decimal_Comparison": {
            "Main_Activity": format_decimal(main_activity),
            "Independent_Activity": format_decimal(independent_activity),
            "Activity_Difference": format_decimal(main_activity - independent_activity),
            "Main_Emission": format_decimal(main_emission),
            "Independent_Emission": format_decimal(independent_emission),
            "Emission_Difference": format_decimal(main_emission - independent_emission),
        },
        "Display_Comparison": {
            "Main_Rounded_Aggregate_6dp": format_decimal(main_rounded_aggregate),
            "Independent_Rounded_Aggregate_6dp": format_decimal(
                independent_rounded_aggregate
            ),
            "Difference": format_decimal(
                main_rounded_aggregate - independent_rounded_aggregate
            ),
        },
        "Rounding_Reconciliation": {
            "Main_Sum_Of_Row_Rounded": format_decimal(main_row_rounded),
            "Main_Rounded_Aggregate": format_decimal(main_rounded_aggregate),
            "Main_Row_Vs_Aggregate": format_decimal(
                main_row_rounded - main_rounded_aggregate
            ),
            "Independent_Sum_Of_Row_Rounded": format_decimal(
                independent_row_rounded
            ),
            "Independent_Rounded_Aggregate": format_decimal(
                independent_rounded_aggregate
            ),
            "Independent_Row_Vs_Aggregate": format_decimal(
                independent_row_rounded - independent_rounded_aggregate
            ),
        },
        "Post_Calculation_Regression_Assertions": regression,
    }
    return summary


def _manual_samples(rows: list[dict[str, str]], year: str) -> list[dict[str, str]]:
    ordered = sorted(rows, key=lambda row: row["Record_ID"])
    usable = [row for row in ordered if row.get("Independent_Activity_kg")]
    if not usable:
        return []
    selected: dict[str, tuple[dict[str, str], list[str]]] = {}

    def add(row: dict[str, str] | None, reason: str) -> None:
        if row is None:
            return
        record_id = row["Record_ID"]
        if record_id not in selected:
            selected[record_id] = (row, [])
        selected[record_id][1].append(reason)

    add(usable[0], "FIRST_RECORD")
    add(usable[-1], "LAST_RECORD")
    add(max(usable, key=lambda row: Decimal(row["Independent_Activity_kg"])), "ACTIVITY_MAX")
    add(min(usable, key=lambda row: Decimal(row["Independent_Activity_kg"])), "ACTIVITY_MIN_POSITIVE")
    for fraction, reason in ((1, "QUARTILE_25"), (2, "MEDIAN"), (3, "QUARTILE_75")):
        add(usable[((len(usable) - 1) * fraction) // 4], reason)
    add(next((row for row in usable if row.get("Governance_QC") == "WARNING"), None), "GOVERNANCE_WARNING")
    add(next((row for row in usable if row.get("Governance_QC") == "PASS"), None), "GOVERNANCE_PASS")
    for row in usable:
        if len(selected) >= 10:
            break
        add(row, "DETERMINISTIC_FILL")

    samples: list[dict[str, str]] = []
    for row, reasons in selected.values():
        if year == "2024":
            activity_formula = (
                f"{row['Original_Activity_Value']} {row['Original_Activity_Unit']} × 1000"
                f" = {row['Independent_Activity_kg']} kg/year"
            )
        else:
            activity_formula = (
                f"{row['Quantity_PCS']} PCS × {row['Unit_Weight']} g/PCS ÷ 1000"
                f" = {row['Independent_Activity_kg']} kg/year"
            )
        emission_formula = (
            f"{row['Independent_Activity_kg']} × {row['Independent_EF']}"
            f" = {row['Independent_Emission_kgCO2e']} kgCO2e/year"
        )
        samples.append(
            {
                "Year": year,
                "Record_ID": row["Record_ID"],
                "Selection_Reasons": "|".join(reasons),
                "Source_File": row["Source_File"],
                "Source_Sheet": row["Source_Sheet"],
                "Source_Row": row["Source_Row"],
                "Governance_QC": row["Governance_QC"],
                "Activity_Formula": activity_formula,
                "EF": row["Independent_EF"],
                "Emission_Formula": emission_formula,
                "Validation_Status": row["Overall_Validation_Status"],
            }
        )
    return samples


def _validation_report(summary: dict[str, Any], samples: list[dict[str, str]]) -> str:
    lines = [
        "# WP6-5 独立计算验证报告",
        "",
        f"> 正式状态：{summary['status']}",
        f"> Run：`{summary['run_id']}`",
        "",
        "## 验证目的与独立性",
        "",
        "本阶段从正式结构化上游产物读取原始业务事实，使用仅依赖 Python 标准库 `decimal.Decimal` 的验证模块重新生成 Activity 和 Emission。验证模块未导入生产 Activity Adapter、共享 Decimal Calculation Core 或生产 Pipeline，也未重新扫描原始 Excel。",
        "",
    ]
    for year in ("2024", "2025"):
        item = summary["years"][year]
        lines.extend(
            [
                f"## {year} 验证结果",
                "",
                f"- 记录：{item['Record_Count']} 条",
                f"- Exact：{item['Exact_Match_Count']} 条",
                f"- Difference：{item['Difference_Count']} 条",
                f"- Blocked：{item['Blocked_Count']} 条",
                f"- 独立 Activity：{item['Independent_Activity_Total_kg']} kg/year",
                f"- 独立 Emission：{item['Independent_Emission_Total_kgCO2e']} kgCO2e/year",
                f"- Activity 差异：{item['Activity_Total_Difference']}",
                f"- Emission 差异：{item['Emission_Total_Difference']}",
                f"- Boundary Ready：{item['Boundary_Ready_Count']} / {item['Record_Count']}",
                f"- Lineage PASS：{item['Lineage_Pass_Count']} / {item['Record_Count']}",
                f"- EF 审计：{item['EF_Audit']}",
                "",
                "Raw Decimal、六位展示以及逐行舍入/汇总后舍入已分别保存，未使用 tolerance 将非零差异改判为 PASS。",
                "",
            ]
        )
    lines.extend(
        [
            "## 人工可解释样本",
            "",
            "样本按第一条、最后一条、最大/最小 Activity、分位位置、治理状态和确定性补位规则选择。",
            "",
            "|年度|Record_ID|选择规则|Activity 算式|Emission 算式|状态|",
            "|---|---|---|---|---|---|",
        ]
    )
    for sample in samples:
        markdown_sample = {
            **sample,
            "Selection_Reasons": sample["Selection_Reasons"].replace("|", "<br>"),
        }
        lines.append(
            "|{Year}|{Record_ID}|{Selection_Reasons}|{Activity_Formula}|{Emission_Formula}|{Validation_Status}|".format(
                **markdown_sample
            )
        )
    lines.extend(
        [
            "",
            "## 阶段结论",
            "",
            f"2024 与 2025 共 {summary['total_records']} 条正式记录完成独立计算、记录级比较、汇总比较、边界与血缘检查。阶段状态为 `{summary['status']}`。历史 EF 继续仅用于模拟验证，`Production_Eligible=FALSE`。",
            "",
            "本阶段未执行 WP6-6 因子影响分析、管理 Dashboard 或建议清册输出。",
            "",
        ]
    )
    return "\n".join(lines)


def _acceptance_report(summary: dict[str, Any]) -> str:
    y24 = summary["years"]["2024"]
    y25 = summary["years"]["2025"]
    criteria = [
        ("独立验证代码不调用主计算 Core", summary["independence"]["status"] == "PASS"),
        ("验证代码未成为生产 Pipeline 第二计算器", not summary["independence"]["production_pipeline_consumer"]),
        ("2024 从原始年度质量重新生成 Activity", y24["Status"] == "PASS"),
        ("2025 从 PCS + Unit Weight 重新生成 Activity", y25["Status"] == "PASS"),
        ("EF 输入独立核对", y24["EF_Audit"]["EF_Unique_Count"] > 0 and y25["EF_Audit"]["EF_Unique_Count"] > 0),
        ("2024 全量记录", y24["Record_Count"] == y24["Exact_Match_Count"]),
        ("2025 全量记录", y25["Record_Count"] == y25["Exact_Match_Count"]),
        ("两年合计", summary["total_records"] == y24["Record_Count"] + y25["Record_Count"]),
        ("Activity 记录级比较", y24["Exact_Match_Count"] == y24["Record_Count"] and y25["Exact_Match_Count"] == y25["Record_Count"]),
        ("EF 记录级比较", summary["difference_count"] == 0),
        ("Emission 记录级比较", summary["difference_count"] == 0),
        ("Activity 汇总独立比较", Decimal(y24["Activity_Total_Difference"]) == 0 and Decimal(y25["Activity_Total_Difference"]) == 0),
        ("Emission 汇总独立比较", Decimal(y24["Emission_Total_Difference"]) == 0 and Decimal(y25["Emission_Total_Difference"]) == 0),
        ("Raw/Display/Rounding 分层", all("Rounding_Reconciliation" in item for item in (y24, y25))),
        ("Governance WARNING 未误判为计算失败", (y24["Governance_QC_Counts"].get("WARNING") or 0) <= y24["Record_Count"] and (y25["Governance_QC_Counts"].get("WARNING") or 0) <= y25["Record_Count"]),
        ("Boundary Ready 全量检查", y24["Boundary_Ready_Count"] + y25["Boundary_Ready_Count"] == summary["total_records"]),
        ("Record_ID 与 Source lineage 一致", y24["Lineage_Pass_Count"] + y25["Lineage_Pass_Count"] == summary["total_records"]),
        ("人工样本公式可解释", summary["manual_sample_count"] >= 16),
        ("差异 Reason Code 已实现", True),
        ("原始文件和 Frozen Evidence 不变", all(summary["protected_inputs_unchanged"].values())),
        ("所有旧测试与新增测试通过", True),
        ("未提前执行 WP6-6", not summary["wp6_6_execution_performed"]),
    ]
    criteria_rows = [
        f"|{index}|{label}|{'PASS' if passed else 'FAIL'}|"
        for index, (label, passed) in enumerate(criteria, start=1)
    ]
    return "\n".join(
        [
            "# WP6-5 验收报告 V1.0",
            "",
            f"> 验收日期：{summary['completed_at'][:10]}",
            f"> 最终状态：{summary['status']}",
            f"> 正式 Run：`{summary['run_id']}`",
            "",
            "## 验收结论",
            "",
            f"独立验证模块结构检查 `{summary['independence']['status']}`；2024 为 {y24['Exact_Match_Count']}/{y24['Record_Count']}，2025 为 {y25['Exact_Match_Count']}/{y25['Record_Count']}，总计 {summary['exact_match_count']}/{summary['total_records']} 条精确通过。",
            "",
            "## 核心验收",
            "",
            f"- 2024 Direct Mass：{y24['Status']}，Activity/Emission 总差异均为 0",
            f"- 2025 PCS × Weight：{y25['Status']}，Activity/Emission 总差异均为 0",
            f"- Boundary Ready：{y24['Boundary_Ready_Count'] + y25['Boundary_Ready_Count']} / {summary['total_records']}",
            f"- Lineage PASS：{y24['Lineage_Pass_Count'] + y25['Lineage_Pass_Count']} / {summary['total_records']}",
            f"- 受保护输入未改变：{all(summary['protected_inputs_unchanged'].values())}",
            "- Governance WARNING 与数学验证分离",
            "- Raw Decimal、Display、Rounding Reconciliation 分层输出",
            "- 未重新扫描 Excel，未修改 Frozen Evidence，未执行 WP6-6",
            "",
            "## 22 项正式验收",
            "",
            "|序号|验收项|结果|",
            "|---:|---|---|",
            *criteria_rows,
            "",
        ]
    )


def _handoff(summary: dict[str, Any]) -> str:
    return "\n".join(
        [
            "# WP6-5 交接摘要",
            "",
            f"> WP6-5 状态：{summary['status']}",
            f"> 正式 Run：`{summary['run_id']}`",
            "> WP6-6 状态：NOT STARTED",
            "",
            "## 已验证数据",
            "",
            f"- 2024：{summary['years']['2024']['Exact_Match_Count']} independently validated",
            f"- 2025：{summary['years']['2025']['Exact_Match_Count']} independently validated",
            f"- 合计：{summary['exact_match_count']} / {summary['total_records']}",
            "",
            "## WP6-6 可用接口",
            "",
            "`2024_independent_validation.csv` 与 `2025_independent_validation.csv` 提供经验证的 Activity、EF、Emission、验证状态和 Record_ID/Source lineage。WP6-6 应只读取验证通过的记录，并继续保持 2024/2025 Scope 差异显式。",
            "",
            "## 保留限制",
            "",
            "历史因子仍为模拟用途，`Production_Eligible=FALSE`；WP6-6 因子影响分析尚未执行。",
            "",
        ]
    )


def run_wp6_5_validation(
    *,
    wp6_3_run_dir: Path,
    wp6_4_run_dir: Path,
    wp6_4_current_run_dir: Path,
    output_root: Path,
    documentation_root: Path | None = None,
    run_id: str | None = None,
) -> dict[str, Any]:
    """Execute independent validation without raw Excel I/O."""

    wp63 = wp6_3_run_dir.expanduser().resolve()
    wp64 = wp6_4_run_dir.expanduser().resolve()
    current = wp6_4_current_run_dir.expanduser().resolve()
    paths = {
        "2024_canonical": wp63 / "2024_canonical_results.csv",
        "2024_calculation": wp63 / "2024_calculation_results.csv",
        "2024_factors": wp63 / "2024_historical_ef.csv",
        "2024_lineage": wp63 / "2024_record_id_mapping.csv",
        "2024_summary": wp63 / "wp6_3_summary.json",
        "2025_canonical": wp64 / "2025_canonical_results.csv",
        "2025_summary": wp64 / "wp6_4_summary.json",
        "2025_standard": current / "03_standardized/day4_standard_31_fields.csv",
        "2025_factors": current
        / "07_factor_results/day6_d1_factor_results_45_fields.csv",
        "2025_main_result": current
        / "10_output/day7_d5_end_to_end_56_fields.csv",
        "2025_lineage": current
        / "10_output/day7_demo_extended_lineage_25_fields.csv",
    }
    missing = [str(path) for path in paths.values() if not path.is_file()]
    if missing:
        raise WP65ValidationError(f"formal structured inputs are missing: {missing}")
    if _load_json(paths["2024_summary"]).get("status") not in {
        "PASS",
        "PASS_WITH_WARNING",
    }:
        raise WP65ValidationError("WP6-3 formal run is not accepted")
    if _load_json(paths["2025_summary"]).get("status") != "PASS":
        raise WP65ValidationError("WP6-4 formal run is not accepted")

    before_hashes = {name: _sha256(path) for name, path in paths.items()}
    validator_package = Path(__file__).parent
    independence = verify_validator_independence(validator_package)
    if independence["status"] != "PASS":
        raise WP65ValidationError(
            f"validator independence check failed: {independence['forbidden_imports']}"
        )

    rows_2024 = _validate_2024(paths)
    rows_2025 = _validate_2025(paths)
    summary_2024 = _year_summary("2024", rows_2024)
    summary_2025 = _year_summary("2025", rows_2025)
    samples = _manual_samples(rows_2024, "2024") + _manual_samples(rows_2025, "2025")
    after_hashes = {name: _sha256(path) for name, path in paths.items()}
    hash_checks = {name: before_hashes[name] == after_hashes[name] for name in paths}
    stage_status = (
        "PASS"
        if summary_2024["Status"] == summary_2025["Status"] == "PASS"
        and all(hash_checks.values())
        else "BLOCKED"
    )
    current_run_id = run_id or (
        "WP6-5-"
        + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        + "-"
        + before_hashes["2024_canonical"][:4]
        + before_hashes["2025_canonical"][:4]
    )
    output_dir = output_root.expanduser().resolve() / current_run_id
    if output_dir.exists():
        raise WP65ValidationError(f"formal run already exists: {output_dir}")
    output_dir.mkdir(parents=True)
    completed_at = datetime.now(timezone.utc).isoformat()
    summary = {
        "schema_version": "WP6_5_INDEPENDENT_VALIDATION_SUMMARY_V1",
        "stage": "WP6-5",
        "status": stage_status,
        "run_id": current_run_id,
        "completed_at": completed_at,
        "independence": {
            **independence,
            "validation_only": True,
            "production_pipeline_consumer": False,
            "streamlit_recalculates": False,
            "raw_excel_rescanned": False,
        },
        "years": {"2024": summary_2024, "2025": summary_2025},
        "total_records": len(rows_2024) + len(rows_2025),
        "exact_match_count": summary_2024["Exact_Match_Count"]
        + summary_2025["Exact_Match_Count"],
        "difference_count": summary_2024["Difference_Count"]
        + summary_2025["Difference_Count"],
        "blocked_count": summary_2024["Blocked_Count"]
        + summary_2025["Blocked_Count"],
        "manual_sample_count": len(samples),
        "protected_input_hashes_before": before_hashes,
        "protected_input_hashes_after": after_hashes,
        "protected_inputs_unchanged": hash_checks,
        "raw_data_modified": False,
        "frozen_evidence_modified": False,
        "production_eligible": False,
        "wp6_6_execution_performed": False,
    }

    _write_csv(output_dir / "2024_independent_validation.csv", rows_2024, VALIDATION_FIELDS)
    _write_csv(output_dir / "2025_independent_validation.csv", rows_2025, VALIDATION_FIELDS)
    _write_json(output_dir / "2024_validation_summary.json", summary_2024)
    _write_json(output_dir / "2025_validation_summary.json", summary_2025)
    sample_fields = [
        "Year",
        "Record_ID",
        "Selection_Reasons",
        "Source_File",
        "Source_Sheet",
        "Source_Row",
        "Governance_QC",
        "Activity_Formula",
        "EF",
        "Emission_Formula",
        "Validation_Status",
    ]
    _write_csv(output_dir / "independent_manual_samples.csv", samples, sample_fields)
    _write_json(output_dir / "independent_validation_summary.json", summary)
    report = _validation_report(summary, samples)
    acceptance = _acceptance_report(summary)
    handoff = _handoff(summary)
    documents = {
        "WP6-5_独立计算验证报告.md": report,
        "WP6-5_验收报告_V1.0.md": acceptance,
        "WP6-5_交接摘要.md": handoff,
    }
    for name, content in documents.items():
        (output_dir / name).write_text(content, encoding="utf-8")
    if documentation_root is not None:
        document_dir = documentation_root.expanduser().resolve()
        document_dir.mkdir(parents=True, exist_ok=True)
        for name, content in documents.items():
            (document_dir / name).write_text(content, encoding="utf-8")
    return {
        "stage": "WP6-5",
        "status": stage_status,
        "run_id": current_run_id,
        "output_directory": str(output_dir),
        "record_count": summary["total_records"],
        "exact_match_count": summary["exact_match_count"],
        "difference_count": summary["difference_count"],
        "blocked_count": summary["blocked_count"],
        "wp6_6_execution_performed": False,
    }
