"""Day 4 conversion to the frozen WP2 31-field standard structure."""

from __future__ import annotations

import csv
import hashlib
import json
import re
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.errors import PipelineUserError
from carbon_excel_pipeline.standardization.mapping_catalog import (
    MappingCatalog,
    find_supplier,
    load_frozen_id_map,
    load_inline_mapping_catalog,
    load_private_mapping_catalog,
)
from carbon_excel_pipeline.standardization.record_ids import RecordIdResolver


TRACE_FIELDS = [
    "Record_ID",
    "Source_File",
    "Source_Sheet",
    "Source_Row",
    "Record_ID_Method",
    "Supplier_Mapping_ID",
    "Supplier_Mapping_Source",
    "Supplier_Confidence",
    "Customer_Mapping_ID",
    "Customer_Mapping_Source",
    "Customer_Confidence",
    "Project_Cell_Mapping_ID",
    "Project_Cell_Mapping_Source",
    "Project_Cell_Confidence",
    "Project_Cell_Manual_Confirmed",
    "Mapping_Trace_Status",
]


def _load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, payload: Any) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def _write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _pipeline_error(
    *, code: str, message: str, location: str, value: Any, rule: str, suggestion: str
) -> PipelineUserError:
    return PipelineUserError(
        stage="STANDARDIZATION_31_FIELDS",
        error_code=code,
        message_cn=message,
        source_location=location,
        original_value=value,
        rule=rule,
        impact="阻断整次Day 4运行",
        fix_suggestion=suggestion,
    )


def _contract_fields(contract: dict[str, Any]) -> list[str]:
    fields = [item["name"] for item in contract["fields"]]
    if contract.get("field_count") != 31 or len(fields) != 31 or len(set(fields)) != 31:
        raise ValueError("The Day 4 contract must contain 31 unique ordered fields.")
    return fields


def _validate_record(
    record: dict[str, Any], contract: dict[str, Any]
) -> list[dict[str, Any]]:
    errors: list[dict[str, Any]] = []
    for spec in contract["fields"]:
        name = spec["name"]
        value = record.get(name)
        if spec.get("required") and (value is None or value == ""):
            errors.append({"field": name, "error_code": "REQUIRED_VALUE_MISSING"})
            continue
        allowed = spec.get("allowed")
        if allowed is not None and value not in allowed:
            errors.append({"field": name, "error_code": "VALUE_NOT_ALLOWED"})
        try:
            if spec["type"] == "integer" and Decimal(str(value)) != Decimal(
                str(value)
            ).to_integral_value():
                errors.append({"field": name, "error_code": "NOT_INTEGER"})
            elif spec["type"] == "decimal" and not Decimal(str(value)).is_finite():
                errors.append({"field": name, "error_code": "NOT_FINITE_DECIMAL"})
        except (InvalidOperation, ValueError):
            errors.append({"field": name, "error_code": "TYPE_VALIDATION_FAILED"})
    return errors


def _unknown_customer(mapping_config: dict[str, Any]) -> dict[str, Any]:
    unknown = mapping_config["missing_values"]["unknown"]
    return {
        "Raw_Customer_Value": unknown,
        "Customer_Name": unknown,
        "Customer_ID": unknown,
        "Customer_Mapping_Status": "UNMAPPED",
        "Customer_Mapping_Source": mapping_config["missing_values"][
            "no_mapping_source"
        ],
        "Mapping_ID": "",
        "Confidence": "NONE",
    }


def _mapped_customer(
    catalog: MappingCatalog,
    customer_id: Any,
    mapping_config: dict[str, Any],
) -> dict[str, Any]:
    if customer_id in (None, "", "UNKNOWN"):
        return _unknown_customer(mapping_config)
    source = catalog.customers_by_id.get(str(customer_id))
    if source is None:
        return _unknown_customer(mapping_config)
    return {
        "Raw_Customer_Value": source["Raw_Customer_Value"],
        "Customer_Name": source["Customer_Name"],
        "Customer_ID": source["Customer_ID"],
        "Customer_Mapping_Status": "CUSTOMER_MAPPED",
        "Customer_Mapping_Source": source.get("Source", "INLINE_MAPPING"),
        "Mapping_ID": source.get("Mapping_ID", ""),
        "Confidence": source.get("Confidence", "UNKNOWN"),
    }


def _standardize_one(
    candidate: dict[str, str],
    *,
    profile: dict[str, Any],
    mapping_config: dict[str, Any],
    catalog: MappingCatalog,
    id_resolver: RecordIdResolver,
) -> tuple[dict[str, Any], dict[str, Any]]:
    constants = mapping_config["constants"]
    missing = mapping_config["missing_values"]
    record_id, record_id_method = id_resolver.resolve(candidate)
    supplier = find_supplier(catalog, candidate["Supplier_Alias_Candidate"])
    if supplier is None:
        supplier = {
            "Supplier_Abbreviation": missing["pending"],
            "Supplier_Name": missing["pending"],
            "Supplier_ID": missing["pending"],
            "Supplier_Status": "PENDING",
            "Mapping_Source": missing["no_mapping_source"],
            "Mapping_ID": "",
            "Confidence": "NONE",
        }
        supplier_mapped = False
    else:
        supplier_mapped = True

    project_key = (
        int(profile["year"]),
        str(supplier["Supplier_ID"]),
        candidate["Product_Description_Raw"],
    )
    project = catalog.project_cells.get(project_key)
    project_mapped = project is not None
    if project is None:
        project = {
            "Customer_ID": missing["unknown"],
            "Project_Code": missing["pending"],
            "Cell_Model": missing["unknown"],
            "Chemistry": missing["unknown"],
            "Mapping_ID": "",
            "Source": missing["no_mapping_source"],
            "Confidence": "NONE",
            "Manual_Confirmed": "NO",
        }
    customer = _mapped_customer(catalog, project["Customer_ID"], mapping_config)

    issue_codes: list[str] = []
    if not supplier_mapped:
        issue_codes.append("SUPPLIER_UNMAPPED")
    if customer["Customer_Mapping_Status"] != "CUSTOMER_MAPPED":
        issue_codes.append("CUSTOMER_UNMAPPED")
    if not project_mapped:
        issue_codes.append("PROJECT_MAPPING_PENDING")
    if project["Cell_Model"] in (None, "", missing["unknown"]):
        issue_codes.append("CELL_MODEL_UNKNOWN")
    if project["Chemistry"] in (None, "", missing["unknown"]):
        issue_codes.append("CHEMISTRY_UNKNOWN")

    standard = {
        "Record_ID": record_id,
        "Year": int(profile["year"]),
        "Source_File": candidate["Source_File"],
        "Source_Sheet": candidate["Source_Sheet"],
        "Source_Row": int(candidate["Source_Row"]),
        "Business_Unit": constants["business_unit"],
        "Activity_Category": constants["activity_category"],
        "Activity_Category_Code": constants["activity_category_code"],
        "Supplier_Abbreviation": supplier["Supplier_Abbreviation"],
        "Supplier_Name": supplier["Supplier_Name"],
        "Supplier_ID": supplier["Supplier_ID"],
        "Supplier_Status": supplier["Supplier_Status"],
        "Mapping_Source": supplier["Mapping_Source"],
        "Customer_Raw_Value": customer["Raw_Customer_Value"],
        "Customer_Name": customer["Customer_Name"],
        "Customer_ID": customer["Customer_ID"],
        "Customer_Mapping_Status": customer["Customer_Mapping_Status"],
        "Customer_Mapping_Source": customer["Customer_Mapping_Source"],
        "Project_Code": project["Project_Code"],
        "Cell_Model": project["Cell_Model"],
        "Material_ID": constants["material_id_default"],
        "Chemistry": project["Chemistry"],
        "Product_Description": candidate["Product_Description_Raw"],
        "PCS": candidate["PCS_Clean"],
        "Unit_Weight_g": candidate["Unit_Weight_Clean"],
        "Original_Activity_Value": candidate["Annual_Activity_Clean"],
        "Original_Activity_Unit": constants["original_activity_unit"],
        "Data_Status": "COMPLETE" if not issue_codes else "PARTIAL",
        "Mapping_Status": "SUPPLIER_MAPPED" if supplier_mapped else "PENDING",
        "QC_Status": "PASS" if not issue_codes else "WARNING",
        "Issue_Code": "NONE" if not issue_codes else ";".join(issue_codes),
    }
    trace = {
        "Record_ID": record_id,
        "Source_File": candidate["Source_File"],
        "Source_Sheet": candidate["Source_Sheet"],
        "Source_Row": int(candidate["Source_Row"]),
        "Record_ID_Method": record_id_method,
        "Supplier_Mapping_ID": supplier.get("Mapping_ID", ""),
        "Supplier_Mapping_Source": supplier.get("Mapping_Source", ""),
        "Supplier_Confidence": supplier.get("Confidence", "UNKNOWN"),
        "Customer_Mapping_ID": customer.get("Mapping_ID", ""),
        "Customer_Mapping_Source": customer["Customer_Mapping_Source"],
        "Customer_Confidence": customer.get("Confidence", "UNKNOWN"),
        "Project_Cell_Mapping_ID": project.get("Mapping_ID", ""),
        "Project_Cell_Mapping_Source": project.get("Source", ""),
        "Project_Cell_Confidence": project.get("Confidence", "UNKNOWN"),
        "Project_Cell_Manual_Confirmed": project.get("Manual_Confirmed", "NO"),
        "Mapping_Trace_Status": "COMPLETE"
        if supplier_mapped and project_mapped
        else "PENDING",
    }
    return standard, trace


def run_day4_standardization(
    run_dir: Path,
    *,
    profile_config_path: Path,
    contract_path: Path,
    mapping_config_path: Path,
    private_id_baseline_path: Path | None = None,
    private_mapping_workbook_path: Path | None = None,
) -> dict[str, Any]:
    run_dir = run_dir.expanduser().resolve()
    day3_path = run_dir / "03_standardized" / "day3_cleaned_candidates.csv"
    day3_summary_path = run_dir / "02_scope_filter" / "day3_scope_summary.json"
    output_dir = run_dir / "03_standardized"
    if not day3_path.is_file() or not day3_summary_path.is_file():
        raise _pipeline_error(
            code="DAY3_OUTPUT_MISSING",
            message="缺少Day 3清洗候选或阶段摘要。",
            location="03_standardized",
            value="day3_cleaned_candidates.csv",
            rule="Day 4只能接续已完成的Day 3运行目录。",
            suggestion="先执行Day 3 scope-clean并确认状态为PASS。",
        )
    if not output_dir.is_dir():
        raise _pipeline_error(
            code="STANDARD_OUTPUT_DIRECTORY_MISSING",
            message="运行目录缺少标准数据阶段目录。",
            location="03_standardized",
            value="03_standardized",
            rule="运行目录必须由受控导入器完整建立。",
            suggestion="重新执行Day 2接收流程。",
        )

    day3_summary = _load_json(day3_summary_path)
    if day3_summary.get("status") != "PASS":
        raise _pipeline_error(
            code="DAY3_NOT_PASSED",
            message="Day 3阶段状态不是PASS。",
            location="day3_scope_summary.json",
            value=day3_summary.get("status"),
            rule="只有通过范围筛选和清洗的候选才能进入31字段标准化。",
            suggestion="修复Day 3验证错误后重新运行。",
        )

    profile = _load_json(profile_config_path)
    contract = _load_json(contract_path)
    mapping_config = _load_json(mapping_config_path)
    fields = _contract_fields(contract)
    if mapping_config["profile_id"] != profile["profile_id"]:
        raise _pipeline_error(
            code="PROFILE_MAPPING_CONFIG_MISMATCH",
            message="Profile与映射配置不一致。",
            location=mapping_config_path.name,
            value=mapping_config["profile_id"],
            rule="私有和公开Profile不得交叉使用映射配置。",
            suggestion="选择与Profile匹配的Day 4映射配置。",
        )

    frozen_ids: dict[tuple[str, str, int], str] = {}
    evidence_hashes: dict[str, str] = {}
    catalog = load_inline_mapping_catalog(mapping_config)

    candidates = _read_csv(day3_path)
    resolver = RecordIdResolver(profile, mapping_config, frozen_ids)
    standardized: list[dict[str, Any]] = []
    traces: list[dict[str, Any]] = []
    validation_errors: list[dict[str, Any]] = []
    for candidate in candidates:
        if candidate.get("Cleaning_Status") != "PASS":
            validation_errors.append(
                {
                    "source_row": candidate.get("Source_Row"),
                    "error_code": "DAY3_CANDIDATE_NOT_CLEAN",
                }
            )
            continue
        try:
            record, trace = _standardize_one(
                candidate,
                profile=profile,
                mapping_config=mapping_config,
                catalog=catalog,
                id_resolver=resolver,
            )
        except (KeyError, TypeError, ValueError) as error:
            validation_errors.append(
                {
                    "source_row": candidate.get("Source_Row"),
                    "error_code": "STANDARDIZATION_FAILED",
                    "detail": str(error),
                }
            )
            continue
        row_errors = _validate_record(record, contract)
        validation_errors.extend(
            {
                "source_row": candidate.get("Source_Row"),
                **item,
            }
            for item in row_errors
        )
        standardized.append(record)
        traces.append(trace)

    ids = [record["Record_ID"] for record in standardized]
    if len(ids) != len(set(ids)):
        validation_errors.append({"error_code": "RECORD_ID_NOT_UNIQUE"})
    if len(standardized) != len(candidates):
        validation_errors.append(
            {
                "error_code": "OUTPUT_RECORD_COUNT_MISMATCH",
                "expected": len(candidates),
                "actual": len(standardized),
            }
        )
    valid_patterns = (profile["record_id_regex"],)
    invalid_ids = [
        record_id
        for record_id in ids
        if not any(re.fullmatch(pattern, record_id) for pattern in valid_patterns)
    ]
    if invalid_ids:
        validation_errors.append(
            {"error_code": "RECORD_ID_PROFILE_REGEX_FAILED", "count": len(invalid_ids)}
        )

    _write_csv(output_dir / "day4_standard_31_fields.csv", standardized, fields)
    _write_csv(output_dir / "day4_mapping_trace.csv", traces, TRACE_FIELDS)
    issue_counts = Counter(record["Issue_Code"] for record in standardized)
    id_method_counts = Counter(trace["Record_ID_Method"] for trace in traces)
    qc_counts = Counter(record["QC_Status"] for record in standardized)
    summary = {
        "run_id": run_dir.name,
        "status": "PASS" if not validation_errors else "FAIL",
        "profile_id": profile["profile_id"],
        "contract_id": contract["contract_id"],
        "mapping_config_id": mapping_config["config_id"],
        "input_records": len(candidates),
        "output_records": len(standardized),
        "field_count": len(fields),
        "field_order": fields,
        "record_ids_nonblank": all(bool(value) for value in ids),
        "record_ids_unique": len(ids) == len(set(ids)),
        "record_id_method_counts": dict(sorted(id_method_counts.items())),
        "qc_status_counts": dict(sorted(qc_counts.items())),
        "issue_code_counts": dict(sorted(issue_counts.items())),
        "mapping_trace_records": len(traces),
        "mapping_trace_complete": all(
            trace["Mapping_Trace_Status"] == "COMPLETE" for trace in traces
        ),
        "evidence_hashes": evidence_hashes,
        "validation_errors": validation_errors,
        "outputs": {
            "standard_data": "03_standardized/day4_standard_31_fields.csv",
            "mapping_trace": "03_standardized/day4_mapping_trace.csv",
        },
    }
    _write_json(output_dir / "day4_standardization_summary.json", summary)
    return {**summary, "run_directory": str(run_dir)}
