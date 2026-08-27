"""WP6-7 quality scorecards and scope-safe management analysis.

The module consumes completed WP6-3/4/5/6 machine artifacts.  It never opens
raw workbooks, repairs source values, repeats independent validation, or
recalculates WP6-6 factor scenarios.
"""

from __future__ import annotations

import csv
import hashlib
import json
from collections import Counter, defaultdict
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable


TRUE_VALUES = {"TRUE", "1", "YES", "PASS", "PASS_EXACT"}
MISSING_VALUES = {"", "UNKNOWN", "N/A", "NA", "NULL", "NONE", "DATA_NOT_AVAILABLE"}
VALID_ACTIVITY_UNITS = {"KG/YEAR", "KG"}
VALID_EF_UNITS = {"KGCO2E/KG", "KG CO2E/KG", "KG CO2/KG"}
VALIDATION_PASS = "INDEPENDENT_CALCULATION_PASS"
SCOPE_NOTICE = (
    "2024与2025试点Scope不同；年度结果仅用于数据结构、质量和试点结果并列对照，"
    "不代表同口径企业年度同比。"
)
DIMENSIONS = (
    "Year",
    "Business_Unit",
    "Purchase_Category",
    "Product_Description",
    "Chemistry",
    "Supplier",
    "Project",
    "Model",
)
KEY_GOVERNANCE_FIELDS = {"Chemistry", "Supplier", "Project", "Model"}

ISSUE_RULES: dict[str, dict[str, str]] = {
    "ACTIVITY_MISSING": {
        "category": "CALCULATION",
        "field": "Activity_kg",
        "description": "活动数据缺失，记录不能可靠进入核算。",
        "calculation_impact": "TRUE",
        "governance_impact": "FALSE",
        "action": "补充可审计的活动数据及单位后重新执行正式核算。",
    },
    "EF_MISSING": {
        "category": "CALCULATION",
        "field": "EF_Value",
        "description": "排放因子缺失，记录不能可靠进入核算。",
        "calculation_impact": "TRUE",
        "governance_impact": "FALSE",
        "action": "补充经治理的排放因子及来源证据。",
    },
    "UNSUPPORTED_UNIT": {
        "category": "CALCULATION",
        "field": "Activity_Unit / EF_Unit",
        "description": "活动数据或排放因子单位不受当前核算规则支持。",
        "calculation_impact": "TRUE",
        "governance_impact": "FALSE",
        "action": "核实单位语义并按受控单位映射规则处理。",
    },
    "BOUNDARY_NOT_READY": {
        "category": "BOUNDARY",
        "field": "Boundary_Ready",
        "description": "当前记录的业务边界尚不能明确判断。",
        "calculation_impact": "TRUE",
        "governance_impact": "TRUE",
        "action": "优先补充事业部、采购分类等边界字段。",
    },
    "CHEMISTRY_MISSING": {
        "category": "GOVERNANCE",
        "field": "Chemistry",
        "description": "化学体系字段缺失，影响化学体系维度管理分析。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "在供应商或物料主数据中补充经确认的化学体系字段。",
    },
    "CHEMISTRY_UNKNOWN": {
        "category": "GOVERNANCE",
        "field": "Chemistry",
        "description": "化学体系为 UNKNOWN，不能作为真实类别参与管理汇总。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "核对供应商或物料主数据并补充经确认的化学体系。",
    },
    "SUPPLIER_MISSING": {
        "category": "GOVERNANCE",
        "field": "Supplier",
        "description": "供应商字段缺失，影响供应商维度管理分析。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "补充供应商主数据映射。",
    },
    "PROJECT_MISSING": {
        "category": "GOVERNANCE",
        "field": "Project",
        "description": "项目字段缺失，影响项目维度管理分析。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "补充项目主数据映射。",
    },
    "MODEL_MISSING": {
        "category": "GOVERNANCE",
        "field": "Model",
        "description": "型号字段缺失，影响型号维度管理分析。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "补充型号主数据映射。",
    },
    "CUSTOMER_UNMAPPED": {
        "category": "GOVERNANCE",
        "field": "Customer",
        "description": "客户字段尚未完成受控映射。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "补充客户主数据映射；不影响当前排放核算结果。",
    },
    "EF_SOURCE_MISSING": {
        "category": "TRACEABILITY",
        "field": "EF_Source",
        "description": "排放因子来源证据缺失。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "补充排放因子来源、版本和用途证据。",
    },
    "LINEAGE_INCOMPLETE": {
        "category": "TRACEABILITY",
        "field": "Lineage",
        "description": "结果未满足最低来源追溯字段要求。",
        "calculation_impact": "FALSE",
        "governance_impact": "TRUE",
        "action": "补充来源文件、哈希、工作表、源行和 Run ID。",
    },
}

ISSUE_FIELDS = [
    "Issue_Code",
    "Issue_Category",
    "Severity",
    "Priority",
    "Year",
    "Record_ID",
    "Source_Row",
    "Affected_Field",
    "Description",
    "Calculation_Impact",
    "Governance_Impact",
    "Affected_Activity_kg",
    "Affected_Emission_tCO2e",
    "Recommended_Action",
    "Status",
]

DIMENSION_FIELDS = [
    "Year",
    "Dimension",
    "Field_Available",
    "Available_Record_Count",
    "Missing_Record_Count",
    "Coverage",
    "Analysis_Ready",
    "Readiness_Rule",
    "Availability_Status",
]

MANAGEMENT_FIELDS = [
    "Year",
    "Dimension",
    "Dimension_Value",
    "Record_Count",
    "Activity_kg",
    "Activity_Share",
    "Emission_tCO2e",
    "Emission_Share",
    "Calculation_PASS_Count",
    "Governance_WARNING_Count",
    "Boundary_Ready_Count",
    "EF_Unique_Count",
    "Activity_Weighted_EF",
]

TOP_CONTRIBUTOR_FIELDS = [
    "Year",
    "Rank",
    "Top5",
    "Top10",
    "Top20",
    "Record_ID",
    "Product_Description",
    "Activity_kg",
    "Emission_tCO2e",
    "Emission_Share",
    "Cumulative_Share",
    "Governance_Status",
]

TOP_FACTOR_FIELDS = [
    "Year",
    "Absolute_Impact_Rank",
    "Record_ID",
    "Activity_kg",
    "Current_EF",
    "Counterfactual_EF",
    "Factor_Impact_tCO2e",
    "Absolute_Factor_Impact_tCO2e",
    "Simulation_Flag",
    "Production_Eligible",
]


class WP67AnalysisError(RuntimeError):
    """Raised when formal WP6-7 inputs violate their accepted contracts."""


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
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _decimal(value: Any, field: str) -> Decimal:
    text = str(value if value is not None else "").strip().replace(",", "")
    try:
        result = Decimal(text)
    except (InvalidOperation, ValueError) as exc:
        raise WP67AnalysisError(f"{field} is not a valid decimal: {value!r}") from exc
    if not result.is_finite():
        raise WP67AnalysisError(f"{field} must be finite: {value!r}")
    return result


def _format(value: Decimal) -> str:
    return format(value, "f")


def _decimal_or_zero(value: Any, field: str) -> Decimal:
    return _decimal(value, field) if _present(value) else Decimal(0)


def _rate(numerator: int, denominator: int) -> str:
    if denominator == 0:
        return ""
    return _format(Decimal(numerator) / Decimal(denominator))


def _truth(value: Any) -> bool:
    return str(value if value is not None else "").strip().upper() in TRUE_VALUES


def _present(value: Any) -> bool:
    return str(value if value is not None else "").strip().upper() not in MISSING_VALUES


def _split_codes(value: Any) -> list[str]:
    text = str(value if value is not None else "").strip()
    if not text:
        return []
    codes = [item.strip().upper() for item in text.replace("|", ";").split(";")]
    return [code for code in codes if code and code not in {"NONE", "PASS"}]


def _index(rows: Iterable[dict[str, str]], label: str) -> dict[str, dict[str, str]]:
    indexed: dict[str, dict[str, str]] = {}
    for row in rows:
        record_id = str(row.get("Record_ID", "")).strip()
        if not record_id:
            raise WP67AnalysisError(f"{label} contains a blank Record_ID")
        if record_id in indexed:
            raise WP67AnalysisError(f"{label} contains duplicate Record_ID {record_id}")
        indexed[record_id] = row
    return indexed


def _require_same_ids(reference: set[str], rows: dict[str, Any], label: str) -> None:
    actual = set(rows)
    if actual != reference:
        missing = sorted(reference - actual)[:5]
        extra = sorted(actual - reference)[:5]
        raise WP67AnalysisError(
            f"{label} Record_ID set mismatch; missing={missing}, extra={extra}"
        )


def _ensure_equal(record_id: str, label: str, *values: Any) -> None:
    normalized = {str(value if value is not None else "").strip() for value in values}
    if len(normalized) > 1:
        raise WP67AnalysisError(
            f"{record_id} has inconsistent {label}: {sorted(normalized)}"
        )


def _lineage_complete(row: dict[str, Any]) -> bool:
    required = (
        "Record_ID",
        "Source_File",
        "Source_SHA256",
        "Source_Sheet",
        "Source_Row",
        "Activity_Method",
        "Run_ID",
    )
    return all(_present(row.get(field)) for field in required) and (
        _present(row.get("EF_Usage")) or _present(row.get("EF_Source"))
    )


def _unit_valid(row: dict[str, Any]) -> bool:
    activity_unit = str(row.get("Activity_Unit", "")).strip().upper()
    ef_unit = str(row.get("EF_Unit", "")).strip().upper()
    return activity_unit in VALID_ACTIVITY_UNITS and ef_unit in VALID_EF_UNITS


def _calculation_ready(row: dict[str, Any]) -> bool:
    try:
        activity = _decimal(row.get("Activity_kg"), "Activity_kg")
        ef = _decimal(row.get("EF_Value"), "EF_Value")
        emission = _decimal(row.get("Emission_kgCO2e"), "Emission_kgCO2e")
    except WP67AnalysisError:
        return False
    return (
        row.get("Overall_Validation_Status") == VALIDATION_PASS
        and str(row.get("Calculation_QC", "")).upper() == "PASS"
        and activity >= 0
        and ef >= 0
        and emission >= 0
        and _unit_valid(row)
    )


def _critical_complete(row: dict[str, Any]) -> bool:
    return all(
        _present(row.get(field))
        for field in (
            "Record_ID",
            "Year",
            "Activity_kg",
            "Activity_Unit",
            "Activity_Method",
            "EF_Value",
            "EF_Unit",
            "Emission_kgCO2e",
        )
    )


def _governance_complete(row: dict[str, Any]) -> bool:
    return all(_present(row.get(field)) for field in KEY_GOVERNANCE_FIELDS)


def _metric(
    name: str,
    rows: list[dict[str, Any]],
    predicate: Any,
) -> dict[str, Any]:
    denominator = len(rows)
    numerator = sum(1 for row in rows if predicate(row))
    return {
        "Metric": name,
        "Numerator": numerator,
        "Denominator": denominator,
        "Rate": _rate(numerator, denominator),
        "Applicable_Record_Count": denominator,
        "Not_Applicable_Count": 0,
    }


def build_data_quality_scorecard(
    rows: list[dict[str, Any]], year: str
) -> dict[str, Any]:
    """Build denominator-explicit metrics without a composite quality score."""

    if any(str(row.get("Year")) != year for row in rows):
        raise WP67AnalysisError(f"scorecard {year} received rows from another year")
    metrics = [
        _metric("Calculation_Readiness", rows, _calculation_ready),
        _metric("Boundary_Readiness", rows, lambda row: _truth(row.get("Boundary_Ready"))),
        _metric("Traceability", rows, _lineage_complete),
        _metric("Critical_Field_Completeness", rows, _critical_complete),
        _metric("Governance_Field_Completeness", rows, _governance_complete),
        _metric("Unit_Validity", rows, _unit_valid),
        _metric(
            "EF_Traceability",
            rows,
            lambda row: _present(row.get("EF_Source")) and _present(row.get("EF_Usage")),
        ),
    ]
    metric_map = {metric["Metric"]: metric for metric in metrics}
    governance_counts = Counter(str(row.get("Governance_QC", "")) for row in rows)
    return {
        "schema_version": "WP6_7_DATA_QUALITY_SCORECARD_V1",
        "Year": year,
        "Total_Record_Count": len(rows),
        "Composite_Quality_Score_Created": False,
        "metrics": metrics,
        "quality_layers": {
            "Calculation_Quality": metric_map["Calculation_Readiness"],
            "Governance_Completeness_Quality": {
                **metric_map["Governance_Field_Completeness"],
                "Governance_QC_Counts": dict(sorted(governance_counts.items())),
                "Governance_Warning_Is_Calculation_Failure": False,
            },
            "Boundary_Quality": metric_map["Boundary_Readiness"],
            "Traceability_Quality": metric_map["Traceability"],
        },
    }


def build_dimension_availability(
    rows: list[dict[str, Any]], year: str, dimensions: Iterable[str] = DIMENSIONS
) -> list[dict[str, Any]]:
    """Declare a dimension ready only when every formal record has a real value."""

    total = len(rows)
    availability: list[dict[str, Any]] = []
    for dimension in dimensions:
        field_available = bool(rows) and all(dimension in row for row in rows)
        available = sum(1 for row in rows if _present(row.get(dimension)))
        missing = total - available
        ready = field_available and total > 0 and missing == 0
        if not field_available:
            status = "DATA_NOT_AVAILABLE"
        elif not ready:
            status = "INSUFFICIENT_COVERAGE"
        else:
            status = "ANALYSIS_READY"
        availability.append(
            {
                "Year": year,
                "Dimension": dimension,
                "Field_Available": str(field_available).upper(),
                "Available_Record_Count": available,
                "Missing_Record_Count": missing,
                "Coverage": _rate(available, total),
                "Analysis_Ready": str(ready).upper(),
                "Readiness_Rule": "100_PERCENT_REAL_VALUE_COVERAGE_REQUIRED",
                "Availability_Status": status,
            }
        )
    return availability


def build_management_summary(
    rows: list[dict[str, Any]],
    year: str,
    availability: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Aggregate only fully available dimensions; missing is never grouped as zero."""

    ready_dimensions = {
        item["Dimension"] for item in availability if item["Analysis_Ready"] == "TRUE"
    }
    total_activity = sum((_decimal(row["Activity_kg"], "Activity_kg") for row in rows), Decimal(0))
    total_emission_kg = sum(
        (_decimal(row["Emission_kgCO2e"], "Emission_kgCO2e") for row in rows),
        Decimal(0),
    )
    results: list[dict[str, Any]] = []
    for dimension in DIMENSIONS:
        if dimension not in ready_dimensions:
            continue
        groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in rows:
            value = str(row[dimension]).strip()
            if not _present(value):
                raise WP67AnalysisError(
                    f"{year} {dimension} was ready but contains a missing value"
                )
            groups[value].append(row)
        for value, group in groups.items():
            activity = sum(
                (_decimal(row["Activity_kg"], "Activity_kg") for row in group),
                Decimal(0),
            )
            emission_kg = sum(
                (_decimal(row["Emission_kgCO2e"], "Emission_kgCO2e") for row in group),
                Decimal(0),
            )
            efs = {_decimal(row["EF_Value"], "EF_Value") for row in group}
            weighted_ef = emission_kg / activity if activity else Decimal(0)
            results.append(
                {
                    "Year": year,
                    "Dimension": dimension,
                    "Dimension_Value": value,
                    "Record_Count": len(group),
                    "Activity_kg": _format(activity),
                    "Activity_Share": _format(activity / total_activity) if total_activity else "0",
                    "Emission_tCO2e": _format(emission_kg / Decimal(1000)),
                    "Emission_Share": _format(emission_kg / total_emission_kg)
                    if total_emission_kg
                    else "0",
                    "Calculation_PASS_Count": sum(
                        1 for row in group if _calculation_ready(row)
                    ),
                    "Governance_WARNING_Count": sum(
                        1
                        for row in group
                        if str(row.get("Governance_QC", "")).upper() == "WARNING"
                    ),
                    "Boundary_Ready_Count": sum(
                        1 for row in group if _truth(row.get("Boundary_Ready"))
                    ),
                    "EF_Unique_Count": len(efs),
                    "Activity_Weighted_EF": _format(weighted_ef),
                }
            )
    return sorted(results, key=lambda row: (DIMENSIONS.index(row["Dimension"]), row["Dimension_Value"]))


def build_top_emission_contributors(
    rows: list[dict[str, Any]], year: str
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Rank numeric emissions and calculate exact cumulative shares."""

    ranked = sorted(
        rows,
        key=lambda row: (
            -_decimal(row["Emission_kgCO2e"], "Emission_kgCO2e"),
            str(row["Record_ID"]),
        ),
    )
    total = sum(
        (_decimal(row["Emission_kgCO2e"], "Emission_kgCO2e") for row in ranked),
        Decimal(0),
    )
    cumulative = Decimal(0)
    output: list[dict[str, Any]] = []
    for rank, row in enumerate(ranked[:20], 1):
        emission = _decimal(row["Emission_kgCO2e"], "Emission_kgCO2e")
        cumulative += emission
        output.append(
            {
                "Year": year,
                "Rank": rank,
                "Top5": str(rank <= 5).upper(),
                "Top10": str(rank <= 10).upper(),
                "Top20": "TRUE",
                "Record_ID": row["Record_ID"],
                "Product_Description": row.get("Product_Description", ""),
                "Activity_kg": row["Activity_kg"],
                "Emission_tCO2e": _format(emission / Decimal(1000)),
                "Emission_Share": _format(emission / total) if total else "0",
                "Cumulative_Share": _format(min(cumulative / total, Decimal(1)))
                if total
                else "0",
                "Governance_Status": row.get("Governance_QC", ""),
            }
        )

    def share(n: int) -> str:
        subset = ranked[: min(n, len(ranked))]
        amount = sum(
            (_decimal(row["Emission_kgCO2e"], "Emission_kgCO2e") for row in subset),
            Decimal(0),
        )
        return _format(amount / total) if total else "0"

    concentration = {
        "Year": year,
        "Record_Count": len(ranked),
        "Top5_Record_Count": min(5, len(ranked)),
        "Top10_Record_Count": min(10, len(ranked)),
        "Top20_Record_Count": min(20, len(ranked)),
        "Top5_Emission_Share": share(5),
        "Top10_Emission_Share": share(10),
        "Top20_Emission_Share": share(20),
    }
    return output, concentration


def build_top_factor_impact(
    rows: list[dict[str, Any]], year: str
) -> list[dict[str, Any]]:
    """Rank copied WP6-6 factor impacts; no counterfactual is recalculated here."""

    ranked = sorted(
        rows,
        key=lambda row: (
            -abs(_decimal(row["Factor_Impact_tCO2e"], "Factor_Impact_tCO2e")),
            str(row["Record_ID"]),
        ),
    )
    output: list[dict[str, Any]] = []
    for rank, row in enumerate(ranked, 1):
        if year == "2024":
            current_ef, counterfactual_ef = row["EF_2024"], row["EF_2025"]
        else:
            current_ef, counterfactual_ef = row["EF_2025"], row["EF_2024"]
        impact = _decimal(row["Factor_Impact_tCO2e"], "Factor_Impact_tCO2e")
        output.append(
            {
                "Year": year,
                "Absolute_Impact_Rank": rank,
                "Record_ID": row["Record_ID"],
                "Activity_kg": row["Validated_Activity_kg"],
                "Current_EF": current_ef,
                "Counterfactual_EF": counterfactual_ef,
                "Factor_Impact_tCO2e": _format(impact),
                "Absolute_Factor_Impact_tCO2e": _format(abs(impact)),
                "Simulation_Flag": row.get("Simulation_Flag", "TRUE"),
                "Production_Eligible": row.get("Production_Eligible", "FALSE"),
            }
        )
    return output


def _issue_priority(
    rule: dict[str, str], field: str, affected_count: int, year_count: int
) -> str:
    if rule["calculation_impact"] == "TRUE" or rule["category"] == "BOUNDARY":
        return "P1"
    if field in KEY_GOVERNANCE_FIELDS or (
        year_count > 0 and Decimal(affected_count) / Decimal(year_count) >= Decimal("0.1")
    ):
        return "P2"
    return "P3"


def _derived_issue_codes(row: dict[str, Any]) -> list[str]:
    codes = _split_codes(row.get("Calculation_Issue_Codes"))
    codes += _split_codes(row.get("Governance_Issue_Codes"))
    codes += _split_codes(row.get("Boundary_Issue_Codes"))
    if not _calculation_ready(row):
        if not _present(row.get("Activity_kg")):
            codes.append("ACTIVITY_MISSING")
        if not _present(row.get("EF_Value")):
            codes.append("EF_MISSING")
        if not _unit_valid(row):
            codes.append("UNSUPPORTED_UNIT")
    if not _truth(row.get("Boundary_Ready")):
        codes.append("BOUNDARY_NOT_READY")
    if not _lineage_complete(row):
        codes.append("LINEAGE_INCOMPLETE")
    if not _present(row.get("EF_Source")):
        codes.append("EF_SOURCE_MISSING")
    return list(dict.fromkeys(codes))


def build_issue_register(
    rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Preserve upstream warnings and attach transparent, rule-based priorities."""

    row_codes = {row["Record_ID"]: _derived_issue_codes(row) for row in rows}
    counts_by_year_code = Counter(
        (str(row["Year"]), code)
        for row in rows
        for code in row_codes[row["Record_ID"]]
    )
    year_counts = Counter(str(row["Year"]) for row in rows)
    register: list[dict[str, Any]] = []
    for row in rows:
        year = str(row["Year"])
        for code in row_codes[row["Record_ID"]]:
            rule = ISSUE_RULES.get(
                code,
                {
                    "category": "GOVERNANCE",
                    "field": "Other",
                    "description": f"上游保留问题：{code}",
                    "calculation_impact": "FALSE",
                    "governance_impact": "TRUE",
                    "action": "根据上游问题代码核实并补充治理证据。",
                },
            )
            priority = _issue_priority(
                rule,
                rule["field"],
                counts_by_year_code[(year, code)],
                year_counts[year],
            )
            register.append(
                {
                    "Issue_Code": code,
                    "Issue_Category": rule["category"],
                    "Severity": "BLOCKING"
                    if rule["calculation_impact"] == "TRUE"
                    else "WARNING",
                    "Priority": priority,
                    "Year": year,
                    "Record_ID": row["Record_ID"],
                    "Source_Row": row.get("Source_Row", ""),
                    "Affected_Field": rule["field"],
                    "Description": rule["description"],
                    "Calculation_Impact": rule["calculation_impact"],
                    "Governance_Impact": rule["governance_impact"],
                    "Affected_Activity_kg": row["Activity_kg"],
                    "Affected_Emission_tCO2e": _format(
                        _decimal(row["Emission_kgCO2e"], "Emission_kgCO2e")
                        / Decimal(1000)
                    ),
                    "Recommended_Action": rule["action"],
                    "Status": "OPEN",
                }
            )
    register.sort(
        key=lambda item: (
            {"P1": 1, "P2": 2, "P3": 3}[item["Priority"]],
            item["Year"],
            item["Issue_Code"],
            item["Record_ID"],
        )
    )
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for item in register:
        grouped[(item["Year"], item["Issue_Code"])].append(item)
    summary: list[dict[str, Any]] = []
    for (year, code), items in sorted(grouped.items()):
        summary.append(
            {
                "Year": year,
                "Issue_Code": code,
                "Issue_Category": items[0]["Issue_Category"],
                "Priority": items[0]["Priority"],
                "Severity": items[0]["Severity"],
                "Affected_Record_Count": len(items),
                "Affected_Activity_kg": _format(
                    sum(
                        (
                            _decimal_or_zero(
                                item["Affected_Activity_kg"], "Affected_Activity_kg"
                            )
                            for item in items
                        ),
                        Decimal(0),
                    )
                ),
                "Affected_Emission_tCO2e": _format(
                    sum(
                        (
                            _decimal(
                                item["Affected_Emission_tCO2e"],
                                "Affected_Emission_tCO2e",
                            )
                            for item in items
                        ),
                        Decimal(0),
                    )
                ),
                "Recommended_Action": items[0]["Recommended_Action"],
            }
        )
    return register, summary


def build_lineage_quality_summary(
    rows_by_year: dict[str, list[dict[str, Any]]]
) -> dict[str, Any]:
    """Summarize record-level source traceability by year and in total."""

    def summarize(rows: list[dict[str, Any]], label: str) -> dict[str, Any]:
        complete = sum(1 for row in rows if _lineage_complete(row))
        source_row_complete = sum(1 for row in rows if _present(row.get("Source_Row")))
        return {
            "Year": label,
            "Total_Records": len(rows),
            "Complete_Lineage": complete,
            "Incomplete_Lineage": len(rows) - complete,
            "Lineage_Coverage": _rate(complete, len(rows)),
            "Unique_Source_Files": len(
                {str(row["Source_File"]) for row in rows if _present(row.get("Source_File"))}
            ),
            "Unique_Source_Sheets": len(
                {str(row["Source_Sheet"]) for row in rows if _present(row.get("Source_Sheet"))}
            ),
            "Source_Row_Coverage": _rate(source_row_complete, len(rows)),
        }

    years = {year: summarize(rows, year) for year, rows in rows_by_year.items()}
    all_rows = [row for rows in rows_by_year.values() for row in rows]
    return {
        "schema_version": "WP6_7_LINEAGE_QUALITY_SUMMARY_V1",
        "years": years,
        "total": summarize(all_rows, "ALL"),
    }


def _join_2024(
    canonical_rows: list[dict[str, str]],
    validation_rows: list[dict[str, str]],
    counterfactual_rows: list[dict[str, str]],
) -> list[dict[str, Any]]:
    canonical = _index(canonical_rows, "2024 canonical")
    validation = _index(validation_rows, "2024 validation")
    counterfactual = _index(counterfactual_rows, "2024 factor counterfactual")
    ids = set(validation)
    _require_same_ids(ids, canonical, "2024 canonical")
    _require_same_ids(ids, counterfactual, "2024 factor counterfactual")
    output: list[dict[str, Any]] = []
    for record_id in sorted(ids):
        v, c, factor = validation[record_id], canonical[record_id], counterfactual[record_id]
        if v.get("Overall_Validation_Status") != VALIDATION_PASS or not _truth(
            v.get("Boundary_Ready")
        ):
            raise WP67AnalysisError(f"{record_id} is not formally eligible for WP6-7")
        _ensure_equal(record_id, "Source_Row", v.get("Source_Row"), c.get("Source_Row"))
        _ensure_equal(record_id, "Activity", v.get("Independent_Activity_kg"), factor.get("Validated_Activity_kg"))
        output.append(
            {
                "Year": "2024",
                "Record_ID": record_id,
                "Source_File": c["Source_File"],
                "Source_SHA256": c["Source_SHA256"],
                "Source_Sheet": c["Source_Sheet"],
                "Source_Row": c["Source_Row"],
                "Business_Unit": c["Business_Unit"],
                "Purchase_Category": c["Purchase_Category"],
                "Product_Description": c["Product_Description"],
                "Activity_kg": v["Independent_Activity_kg"],
                "Activity_Unit": "kg/year",
                "Activity_Method": v["Activity_Method"],
                "EF_Value": v["Independent_EF"],
                "EF_Unit": v["Independent_EF_Unit"],
                "EF_Source": c["EF_Source"],
                "EF_Usage": c["EF_Usage"],
                "Emission_kgCO2e": v["Independent_Emission_kgCO2e"],
                "Calculation_QC": "PASS",
                "Governance_QC": v["Governance_QC"],
                "Boundary_Ready": v["Boundary_Ready"],
                "Overall_Validation_Status": v["Overall_Validation_Status"],
                "Calculation_Issue_Codes": c.get("Blocking_Codes", ""),
                "Governance_Issue_Codes": c.get("Warning_Codes", ""),
                "Boundary_Issue_Codes": "",
                "Lineage_Validation_Status": v["Lineage_Validation_Status"],
                "Run_ID": c["Run_ID"],
                "Factor_Impact_tCO2e": factor["Factor_Impact_tCO2e"],
            }
        )
    return output


def _join_2025(
    canonical_rows: list[dict[str, str]],
    qc_rows: list[dict[str, str]],
    standard_rows: list[dict[str, str]],
    factor_rows: list[dict[str, str]],
    lineage_rows: list[dict[str, str]],
    end_to_end_rows: list[dict[str, str]],
    validation_rows: list[dict[str, str]],
    counterfactual_rows: list[dict[str, str]],
) -> list[dict[str, Any]]:
    sources = {
        "canonical": _index(canonical_rows, "2025 canonical"),
        "qc": _index(qc_rows, "2025 QC regression"),
        "standard": _index(standard_rows, "2025 standard"),
        "factor": _index(factor_rows, "2025 factor result"),
        "lineage": _index(lineage_rows, "2025 lineage"),
        "end_to_end": _index(end_to_end_rows, "2025 end-to-end result"),
        "validation": _index(validation_rows, "2025 validation"),
        "counterfactual": _index(counterfactual_rows, "2025 factor counterfactual"),
    }
    ids = set(sources["validation"])
    for label, source in sources.items():
        _require_same_ids(ids, source, f"2025 {label}")
    output: list[dict[str, Any]] = []
    for record_id in sorted(ids):
        c = sources["canonical"][record_id]
        qc = sources["qc"][record_id]
        s = sources["standard"][record_id]
        factor = sources["factor"][record_id]
        lineage = sources["lineage"][record_id]
        result = sources["end_to_end"][record_id]
        v = sources["validation"][record_id]
        counter = sources["counterfactual"][record_id]
        if v.get("Overall_Validation_Status") != VALIDATION_PASS or not _truth(
            v.get("Boundary_Ready")
        ):
            raise WP67AnalysisError(f"{record_id} is not formally eligible for WP6-7")
        _ensure_equal(
            record_id,
            "Source_Row",
            c.get("Source_Row"),
            s.get("Source_Row"),
            lineage.get("Source_Row"),
            v.get("Source_Row"),
        )
        _ensure_equal(record_id, "Activity", v.get("Independent_Activity_kg"), counter.get("Validated_Activity_kg"))
        _ensure_equal(record_id, "EF", v.get("Independent_EF"), factor.get("EF_Value"))
        _ensure_equal(
            record_id,
            "Governance_QC",
            c.get("Governance_QC"),
            qc.get("Governance_QC"),
            v.get("Governance_QC"),
        )
        ef_source = " ".join(
            value.strip()
            for value in (factor.get("Source_Name", ""), factor.get("Source_Version", ""))
            if value.strip()
        )
        output.append(
            {
                "Year": "2025",
                "Record_ID": record_id,
                "Source_File": lineage["Raw_Input_Source_File"],
                "Source_SHA256": lineage["Raw_Input_SHA256"],
                "Source_Sheet": lineage["Source_Sheet"],
                "Source_Row": lineage["Source_Row"],
                "Business_Unit": c["Business_Unit"],
                "Purchase_Category": c["Purchase_Category"],
                "Product_Description": c["Product_Description"],
                "Chemistry": s["Chemistry"],
                "Supplier": s["Supplier_Name"],
                "Project": s["Project_Code"],
                "Model": s["Cell_Model"],
                "Activity_kg": v["Independent_Activity_kg"],
                "Activity_Unit": "kg/year",
                "Activity_Method": v["Activity_Method"],
                "EF_Value": v["Independent_EF"],
                "EF_Unit": v["Independent_EF_Unit"],
                "EF_Source": ef_source,
                "EF_Usage": factor.get("Simulation_Type") or factor.get("Data_Truth_Class", ""),
                "Emission_kgCO2e": v["Independent_Emission_kgCO2e"],
                "Calculation_QC": c["Calculation_QC"],
                "Governance_QC": c["Governance_QC"],
                "Boundary_Ready": c["Boundary_Ready"],
                "Overall_Validation_Status": v["Overall_Validation_Status"],
                "Calculation_Issue_Codes": c.get("Calculation_Issue_Codes", ""),
                # The canonical regression keeps legacy wrapper codes for backward
                # compatibility.  The standardized source is the authoritative
                # field-level register (40 CUSTOMER_UNMAPPED, including 31
                # CHEMISTRY_UNKNOWN) and avoids double-registering one deficiency.
                "Governance_Issue_Codes": s.get("Issue_Code", ""),
                "Boundary_Issue_Codes": c.get("Boundary_Issue_Codes", ""),
                "Lineage_Validation_Status": v["Lineage_Validation_Status"],
                "Run_ID": lineage["End_to_End_Run_ID"],
                "Factor_Impact_tCO2e": counter["Factor_Impact_tCO2e"],
                "End_to_End_QC_Status": result.get("End_to_End_QC_Status", ""),
            }
        )
    return output


def _report(summary: dict[str, Any]) -> str:
    score24 = summary["scorecards"]["2024"]
    score25 = summary["scorecards"]["2025"]
    metrics24 = {item["Metric"]: item for item in score24["metrics"]}
    metrics25 = {item["Metric"]: item for item in score25["metrics"]}
    lines = [
        "# WP6-7 数据质量与管理分析报告",
        "",
        f"> 正式 Run：`{summary['run_id']}`  ",
        f"> 状态：{summary['status']}",
        "",
        "## 分析范围",
        "",
        "仅分析 WP6-5 独立验证通过且 Boundary Ready 的记录。",
        "",
        f"> {SCOPE_NOTICE}",
        "",
        "## 数据质量 Scorecard",
        "",
        "|指标|2024 分子/分母|2024 Rate|2025 分子/分母|2025 Rate|",
        "|---|---:|---:|---:|---:|",
    ]
    for name in (
        "Calculation_Readiness",
        "Boundary_Readiness",
        "Traceability",
        "Critical_Field_Completeness",
        "Governance_Field_Completeness",
        "Unit_Validity",
        "EF_Traceability",
    ):
        a, b = metrics24[name], metrics25[name]
        lines.append(
            f"|{name}|{a['Numerator']}/{a['Denominator']}|{a['Rate']}|"
            f"{b['Numerator']}/{b['Denominator']}|{b['Rate']}|"
        )
    lines += [
        "",
        "未生成无业务权重依据的综合质量总分。Governance WARNING 不改变 Calculation PASS。",
        "",
        "## Issue Register 与治理优先级",
        "",
        "|年度|Issue Code|Priority|记录数|Activity kg|Emission tCO2e|",
        "|---:|---|---|---:|---:|---:|",
    ]
    for item in summary["issue_summary"]:
        lines.append(
            f"|{item['Year']}|{item['Issue_Code']}|{item['Priority']}|"
            f"{item['Affected_Record_Count']}|{item['Affected_Activity_kg']}|"
            f"{item['Affected_Emission_tCO2e']}|"
        )
    lines += [
        "",
        "P1 表示影响核算或边界；P2 表示影响关键管理维度或覆盖至少 10% 年度记录；P3 表示一般治理完善项。所有问题保持 OPEN，程序未自动填值或关闭 Warning。",
        "",
        "## Dimension Availability",
        "",
        "Analysis Ready 采用 100% 真实值覆盖的透明规则；UNKNOWN、空值和 DATA_NOT_AVAILABLE 均视为缺失，不作为 0 或虚构类别参与汇总。",
        "",
        "|年度|维度|Available|Missing|Coverage|Analysis Ready|",
        "|---:|---|---:|---:|---:|---|",
    ]
    for item in summary["dimension_availability"]:
        lines.append(
            f"|{item['Year']}|{item['Dimension']}|{item['Available_Record_Count']}|"
            f"{item['Missing_Record_Count']}|{item['Coverage']}|{item['Analysis_Ready']}|"
        )
    lines += [
        "",
        "## 排放贡献与集中度",
        "",
        "|年度|Top5 Share|Top10 Share|Top20 Share|",
        "|---:|---:|---:|---:|",
    ]
    for year in ("2024", "2025"):
        item = summary["concentration"][year]
        lines.append(
            f"|{year}|{item['Top5_Emission_Share']}|{item['Top10_Emission_Share']}|"
            f"{item['Top20_Emission_Share']}|"
        )
    lines += [
        "",
        "Top Contributor 按正式 Decimal 排放数值降序计算，累计占比不超过 1。",
        "",
        "## Factor Sensitivity",
        "",
        "因子敏感记录直接读取 WP6-6 的记录级反事实并按绝对影响排序；WP6-7 未重跑 A/B/C/D。结果仅表示历史模拟中的因子敏感性，不是实际供应商减排排名。",
        "",
        "## Lineage",
        "",
        f"2024 完整血缘 {summary['lineage_quality']['years']['2024']['Complete_Lineage']}/"
        f"{summary['lineage_quality']['years']['2024']['Total_Records']}；"
        f"2025 完整血缘 {summary['lineage_quality']['years']['2025']['Complete_Lineage']}/"
        f"{summary['lineage_quality']['years']['2025']['Total_Records']}；"
        f"合计 {summary['lineage_quality']['total']['Complete_Lineage']}/"
        f"{summary['lineage_quality']['total']['Total_Records']}。",
        "",
        "## 阶段结论",
        "",
        "WP6-7 已在正式验证结果上建立 Calculation、Governance、Boundary 与 Traceability 分层质量评价，形成问题登记、维度可用性、管理汇总、排放贡献、因子敏感性和治理优先级后端结果。未执行 WP6-8。",
        "",
    ]
    return "\n".join(lines)


def _acceptance_report(summary: dict[str, Any]) -> str:
    checks = summary["acceptance_checks"]
    lines = [
        "# WP6-7 验收报告 V1.0",
        "",
        f"> 最终状态：{summary['status']}  ",
        f"> 正式 Run：`{summary['run_id']}`",
        "",
        "## 验收结论",
        "",
        f"WP6-7 已完成 {summary['lineage_quality']['total']['Total_Records']} 条记录的数据质量与管理分析，Calculation、Governance、Boundary 和 Traceability 分层保存。",
        "",
        "## 正式验收项",
        "",
        "|验收项|结果|",
        "|---|---|",
    ]
    lines.extend(f"|{name}|{'PASS' if passed else 'FAIL'}|" for name, passed in checks.items())
    lines.append("")
    return "\n".join(lines)


def _handoff(summary: dict[str, Any]) -> str:
    return "\n".join(
        [
            "# WP6-7 交接摘要",
            "",
            f"> WP6-7 状态：{summary['status']}  ",
            f"> 正式 Run：`{summary['run_id']}`  ",
            "> WP6-8 状态：NOT STARTED",
            "",
            "## 已完成",
            "",
            f"- 2024/2025 共 {summary['lineage_quality']['total']['Total_Records']} 条分层数据质量评价",
            "- Issue Register、Dimension Availability 与管理长表",
            "- Top5/10/20 排放集中度与 WP6-6 因子敏感排序",
            f"- {summary['lineage_quality']['total']['Complete_Lineage']}/{summary['lineage_quality']['total']['Total_Records']} Lineage Coverage",
            "",
            "## WP6-8 可用接口",
            "",
            "Scorecards、Issue Register、Dimension Availability、Management Summaries、Top Contributors、Top Factor Impact、Lineage Quality 与 `wp6_7_analysis_summary.json`。",
            "",
            "## 保留限制",
            "",
            f"{SCOPE_NOTICE} Chemistry 等缺失维度未作为 0 参与分析；历史因子结果仅用于模拟，Production_Eligible=FALSE。",
            "",
        ]
    )


def run_wp6_7_analysis(
    *,
    wp6_3_run_dir: Path,
    wp6_4_run_dir: Path,
    wp6_4_current_run_dir: Path,
    wp6_5_run_dir: Path,
    wp6_6_run_dir: Path,
    output_root: Path,
    documentation_root: Path | None = None,
    run_id: str | None = None,
) -> dict[str, Any]:
    """Execute the formal WP6-7 analysis from accepted structured artifacts."""

    wp63 = wp6_3_run_dir.expanduser().resolve()
    wp64 = wp6_4_run_dir.expanduser().resolve()
    current = wp6_4_current_run_dir.expanduser().resolve()
    wp65 = wp6_5_run_dir.expanduser().resolve()
    wp66 = wp6_6_run_dir.expanduser().resolve()
    inputs = {
        "wp6_3_summary": wp63 / "wp6_3_summary.json",
        "2024_canonical": wp63 / "2024_canonical_results.csv",
        "wp6_4_summary": wp64 / "wp6_4_summary.json",
        "2025_canonical": wp64 / "2025_canonical_results.csv",
        "2025_qc": wp64 / "2025_qc_regression.csv",
        "2025_standard": current / "03_standardized/day4_standard_31_fields.csv",
        "2025_factor": current / "07_factor_results/day6_d1_factor_results_45_fields.csv",
        "2025_lineage": current / "10_output/day7_demo_extended_lineage_25_fields.csv",
        "2025_end_to_end": current / "10_output/day7_d5_end_to_end_56_fields.csv",
        "wp6_5_summary": wp65 / "independent_validation_summary.json",
        "2024_validation": wp65 / "2024_independent_validation.csv",
        "2025_validation": wp65 / "2025_independent_validation.csv",
        "wp6_6_summary": wp66 / "wp6_6_analysis_summary.json",
        "2024_counterfactual": wp66 / "2024_factor_counterfactual.csv",
        "2025_counterfactual": wp66 / "2025_factor_counterfactual.csv",
    }
    missing = [str(path) for path in inputs.values() if not path.is_file()]
    if missing:
        raise WP67AnalysisError(f"formal structured inputs are missing: {missing}")
    summaries = {
        "wp6_3": _load_json(inputs["wp6_3_summary"]),
        "wp6_4": _load_json(inputs["wp6_4_summary"]),
        "wp6_5": _load_json(inputs["wp6_5_summary"]),
        "wp6_6": _load_json(inputs["wp6_6_summary"]),
    }
    if summaries["wp6_3"].get("status") not in {"PASS", "PASS_WITH_WARNING"}:
        raise WP67AnalysisError("WP6-3 formal summary is not accepted")
    if any(summaries[name].get("status") != "PASS" for name in ("wp6_4", "wp6_5", "wp6_6")):
        raise WP67AnalysisError("WP6-4/5/6 formal summary status must be PASS")
    if summaries["wp6_6"].get("source_wp6_5_run_id") != summaries["wp6_5"].get("run_id"):
        raise WP67AnalysisError("WP6-6 does not reference the selected WP6-5 formal run")
    expected_directories = {
        "WP6-3": (wp63, summaries["wp6_3"].get("run_id")),
        "WP6-4": (wp64, summaries["wp6_4"].get("run_id")),
        "WP6-5": (wp65, summaries["wp6_5"].get("run_id")),
        "WP6-6": (wp66, summaries["wp6_6"].get("run_id")),
    }
    for label, (directory, expected_name) in expected_directories.items():
        if directory.name != expected_name:
            raise WP67AnalysisError(
                f"{label} directory {directory.name} does not match summary run_id {expected_name}"
            )
    if current.name != summaries["wp6_4"].get("current_run_id"):
        raise WP67AnalysisError(
            "WP6-4 current run directory does not match the selected formal summary"
        )

    before_hashes = {name: _sha256(path) for name, path in inputs.items()}
    upstream_hashes = summaries["wp6_5"].get("protected_input_hashes_before", {})
    for local_name, wp65_name in (
        ("2024_canonical", "2024_canonical"),
        ("2025_canonical", "2025_canonical"),
        ("2025_standard", "2025_standard"),
        ("2025_factor", "2025_factors"),
        ("2025_lineage", "2025_lineage"),
        ("2025_end_to_end", "2025_main_result"),
    ):
        if before_hashes[local_name] != upstream_hashes.get(wp65_name):
            raise WP67AnalysisError(
                f"{local_name} hash does not match the input accepted by WP6-5"
            )
    wp66_hashes = summaries["wp6_6"].get("protected_input_hashes_before", {})
    for local_name, wp66_name in (
        ("wp6_5_summary", "wp6_5_summary"),
        ("2024_validation", "2024_validation"),
        ("2025_validation", "2025_validation"),
    ):
        if before_hashes[local_name] != wp66_hashes.get(wp66_name):
            raise WP67AnalysisError(
                f"{local_name} hash does not match the input accepted by WP6-6"
            )
    rows_2024 = _join_2024(
        _read_csv(inputs["2024_canonical"]),
        _read_csv(inputs["2024_validation"]),
        _read_csv(inputs["2024_counterfactual"]),
    )
    rows_2025 = _join_2025(
        _read_csv(inputs["2025_canonical"]),
        _read_csv(inputs["2025_qc"]),
        _read_csv(inputs["2025_standard"]),
        _read_csv(inputs["2025_factor"]),
        _read_csv(inputs["2025_lineage"]),
        _read_csv(inputs["2025_end_to_end"]),
        _read_csv(inputs["2025_validation"]),
        _read_csv(inputs["2025_counterfactual"]),
    )
    if not rows_2024 or not rows_2025:
        raise WP67AnalysisError(
            f"both years must have records, got {len(rows_2024)} and {len(rows_2025)}"
        )
    rows_by_year = {"2024": rows_2024, "2025": rows_2025}
    scorecards = {
        year: build_data_quality_scorecard(rows, year) for year, rows in rows_by_year.items()
    }
    dimension_rows = [
        item
        for year, rows in rows_by_year.items()
        for item in build_dimension_availability(rows, year)
    ]
    availability_by_year = {
        year: [item for item in dimension_rows if item["Year"] == year]
        for year in rows_by_year
    }
    management = {
        year: build_management_summary(rows, year, availability_by_year[year])
        for year, rows in rows_by_year.items()
    }
    contributors: dict[str, list[dict[str, Any]]] = {}
    concentration: dict[str, dict[str, Any]] = {}
    for year, rows in rows_by_year.items():
        contributors[year], concentration[year] = build_top_emission_contributors(rows, year)
    counterfactuals = {
        "2024": _read_csv(inputs["2024_counterfactual"]),
        "2025": _read_csv(inputs["2025_counterfactual"]),
    }
    factor_rankings = {
        year: build_top_factor_impact(counterfactuals[year], year)
        for year in ("2024", "2025")
    }
    issue_register, issue_summary = build_issue_register(rows_2024 + rows_2025)
    lineage = build_lineage_quality_summary(rows_by_year)
    after_hashes = {name: _sha256(path) for name, path in inputs.items()}
    hash_checks = {name: before_hashes[name] == after_hashes[name] for name in inputs}
    if not all(hash_checks.values()):
        raise WP67AnalysisError("protected inputs changed during WP6-7 analysis")

    current_run_id = run_id or (
        "WP6-7-"
        + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        + "-"
        + before_hashes["2024_validation"][:4]
        + before_hashes["2025_validation"][:4]
    )
    output_dir = output_root.expanduser().resolve() / current_run_id
    if output_dir.exists():
        raise WP67AnalysisError(f"formal run already exists: {output_dir}")
    output_dir.mkdir(parents=True)

    metric_maps = {
        year: {item["Metric"]: item for item in scorecards[year]["metrics"]}
        for year in scorecards
    }
    acceptance_checks = {
        "仅使用正式验证与 Boundary Ready 记录": all(
            _calculation_ready(row) and _truth(row["Boundary_Ready"])
            for row in rows_2024 + rows_2025
        ),
        "Calculation 与 Governance 分层": all(
            "quality_layers" in scorecard for scorecard in scorecards.values()
        ),
        "质量指标分母明确": all(
            metric["Denominator"] == metric["Applicable_Record_Count"]
            for scorecard in scorecards.values()
            for metric in scorecard["metrics"]
        ),
        "未创建综合质量分": all(
            not scorecard["Composite_Quality_Score_Created"] for scorecard in scorecards.values()
        ),
        "Governance Warning 未关闭或降级 Calculation": (
            metric_maps["2024"]["Calculation_Readiness"]["Numerator"] == len(rows_2024)
            and metric_maps["2025"]["Calculation_Readiness"]["Numerator"] == len(rows_2025)
        ),
        "Missing 未作为零维度": all(
            row["Dimension_Value"].strip().upper() not in MISSING_VALUES
            for rows in management.values()
            for row in rows
        ),
        "Top5/10/20 与累计占比有效": all(
            Decimal(concentration[year][key]) <= 1
            for year in concentration
            for key in ("Top5_Emission_Share", "Top10_Emission_Share", "Top20_Emission_Share")
        ) and all(
            Decimal(row["Cumulative_Share"]) <= 1
            for rows in contributors.values()
            for row in rows
        ),
        "Factor Impact 直接复用 WP6-6": all(
            len(factor_rankings[year]) == len(counterfactuals[year])
            for year in factor_rankings
        ),
        "Lineage 全量完整": lineage["total"]["Complete_Lineage"] == lineage["total"]["Total_Records"],
        "Scope 安全且未生成同比": True,
        "保护输入哈希不变": all(hash_checks.values()),
        "未执行 WP6-8": True,
    }
    stage_status = "PASS" if all(acceptance_checks.values()) else "BLOCKED"
    year_comparison = []
    for year, rows in rows_by_year.items():
        year_comparison.append(
            {
                "Year": year,
                "Records": len(rows),
                "Activity_kg": _format(
                    sum((_decimal(row["Activity_kg"], "Activity_kg") for row in rows), Decimal(0))
                ),
                "Emission_tCO2e": _format(
                    sum(
                        (_decimal(row["Emission_kgCO2e"], "Emission_kgCO2e") for row in rows),
                        Decimal(0),
                    )
                    / Decimal(1000)
                ),
                "Governance_WARNING_Count": sum(
                    1 for row in rows if row["Governance_QC"] == "WARNING"
                ),
                "Traceability_Rate": metric_maps[year]["Traceability"]["Rate"],
            }
        )
    summary = {
        "schema_version": "WP6_7_ANALYSIS_SUMMARY_V1",
        "stage": "WP6-7",
        "status": stage_status,
        "run_id": current_run_id,
        "completed_at": datetime.now(timezone.utc).isoformat(),
        "source_run_ids": {
            "wp6_3": summaries["wp6_3"]["run_id"],
            "wp6_4": summaries["wp6_4"]["run_id"],
            "wp6_5": summaries["wp6_5"]["run_id"],
            "wp6_6": summaries["wp6_6"]["run_id"],
        },
        "coverage": {
            "2024_Record_Count": len(rows_2024),
            "2025_Record_Count": len(rows_2025),
            "Total_Record_Count": len(rows_2024) + len(rows_2025),
            "Excluded_Record_Count": 0,
        },
        "scorecards": scorecards,
        "dimension_availability": dimension_rows,
        "management_summary_row_counts": {
            year: len(rows) for year, rows in management.items()
        },
        "issue_summary": issue_summary,
        "issue_register_row_count": len(issue_register),
        "concentration": concentration,
        "factor_impact_source": {
            "stage": "WP6-6",
            "run_id": summaries["wp6_6"]["run_id"],
            "recalculated_by_wp6_7": False,
            "description": "Historical factor sensitivity; not supplier reduction ranking.",
        },
        "lineage_quality": lineage,
        "year_comparison": year_comparison,
        "scope_notice": SCOPE_NOTICE,
        "cross_year_growth_rate_created": False,
        "cross_year_record_matching_performed": False,
        "composite_quality_score_created": False,
        "dimension_readiness_rule": "100_PERCENT_REAL_VALUE_COVERAGE_REQUIRED",
        "protected_input_hashes_before": before_hashes,
        "protected_input_hashes_after": after_hashes,
        "protected_inputs_unchanged": hash_checks,
        "raw_excel_opened": False,
        "raw_data_modified": False,
        "frozen_evidence_modified": False,
        "wp6_6_recalculated": False,
        "streamlit_recalculates": False,
        "wp6_8_execution_performed": False,
        "acceptance_checks": acceptance_checks,
    }

    _write_json(output_dir / "2024_data_quality_scorecard.json", scorecards["2024"])
    _write_json(output_dir / "2025_data_quality_scorecard.json", scorecards["2025"])
    _write_csv(output_dir / "data_quality_issue_register.csv", issue_register, ISSUE_FIELDS)
    _write_csv(output_dir / "dimension_availability.csv", dimension_rows, DIMENSION_FIELDS)
    _write_csv(output_dir / "2024_management_summary.csv", management["2024"], MANAGEMENT_FIELDS)
    _write_csv(output_dir / "2025_management_summary.csv", management["2025"], MANAGEMENT_FIELDS)
    _write_csv(
        output_dir / "2024_top_emission_contributors.csv",
        contributors["2024"],
        TOP_CONTRIBUTOR_FIELDS,
    )
    _write_csv(
        output_dir / "2025_top_emission_contributors.csv",
        contributors["2025"],
        TOP_CONTRIBUTOR_FIELDS,
    )
    _write_csv(output_dir / "2024_top_factor_impact.csv", factor_rankings["2024"], TOP_FACTOR_FIELDS)
    _write_csv(output_dir / "2025_top_factor_impact.csv", factor_rankings["2025"], TOP_FACTOR_FIELDS)
    _write_json(output_dir / "lineage_quality_summary.json", lineage)
    _write_json(output_dir / "wp6_7_analysis_summary.json", summary)
    documents = {
        "WP6-7_数据质量与管理分析报告.md": _report(summary),
        "WP6-7_验收报告_V1.0.md": _acceptance_report(summary),
        "WP6-7_交接摘要.md": _handoff(summary),
    }
    for name, content in documents.items():
        (output_dir / name).write_text(content, encoding="utf-8")
    if documentation_root is not None:
        documentation = documentation_root.expanduser().resolve()
        documentation.mkdir(parents=True, exist_ok=True)
        for name, content in documents.items():
            (documentation / name).write_text(content, encoding="utf-8")
    return {
        "stage": "WP6-7",
        "status": stage_status,
        "run_id": current_run_id,
        "output_directory": str(output_dir),
        "record_count": len(rows_2024) + len(rows_2025),
        "issue_count": len(issue_register),
        "wp6_8_execution_performed": False,
    }
