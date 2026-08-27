"""Confirm Day 2 automatic mappings before scope filtering."""

from __future__ import annotations

from typing import Any


def confirm_sheet_mapping(
    mapping_reports: list[dict[str, Any]],
    *,
    target_sheet: str,
    required_targets: list[str],
    mapping_overrides: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    sheet_report = next(
        (item for item in mapping_reports if item.get("sheet_name") == target_sheet), None
    )
    if sheet_report is None:
        return {
            "status": "MANUAL_CONFIRMATION_REQUIRED",
            "sheet_name": target_sheet,
            "requires_manual_confirmation": True,
            "confirmed_fields": [],
            "errors": [
                {
                    "error_code": "TARGET_SHEET_MAPPING_NOT_FOUND",
                    "message_cn": "没有找到目标工作表的字段映射结果。",
                    "fix_suggestion": "重新运行Day 2结构识别，或选择正确的目标工作表。",
                }
            ],
        }

    mapping_by_target = {
        item["target_field"]: item for item in sheet_report.get("mapping_preview", [])
    }
    overrides = mapping_overrides or {}
    confirmed_fields: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    for target in required_targets:
        override = overrides.get(target)
        if override is not None:
            try:
                column_index = int(override["column_index"])
                source_header = str(override["source_header"])
                column_letter = str(override["column_letter"])
            except (KeyError, TypeError, ValueError):
                errors.append(
                    {
                        "error_code": "MANUAL_MAPPING_INVALID",
                        "target_field": target,
                        "message_cn": "人工字段映射缺少有效的源列信息。",
                        "fix_suggestion": "重新选择源列并确认列序号、列字母和表头。",
                    }
                )
                continue
            if column_index < 1 or not source_header or not column_letter:
                errors.append(
                    {
                        "error_code": "MANUAL_MAPPING_INVALID",
                        "target_field": target,
                        "message_cn": "人工字段映射包含空表头或无效列号。",
                        "fix_suggestion": "为目标字段选择一个非空且唯一的源列。",
                    }
                )
                continue
            confirmed_fields.append(
                {
                    "target_field": target,
                    "source_header": source_header,
                    "column_index": column_index,
                    "column_letter": column_letter,
                    "confirmation_method": "USER_CONFIRMED_OVERRIDE",
                    "confirmation_status": "CONFIRMED",
                }
            )
            continue
        item = mapping_by_target.get(target)
        if item is None or item.get("status") != "AUTO_MATCHED":
            errors.append(
                {
                    "error_code": "TARGET_FIELD_NOT_UNIQUELY_MAPPED",
                    "target_field": target,
                    "observed_status": item.get("status") if item else "MISSING",
                    "message_cn": "目标字段没有唯一自动映射。",
                    "fix_suggestion": "人工选择唯一源列后再继续。",
                }
            )
            continue
        columns = item.get("matched_columns", [])
        if len(columns) != 1:
            errors.append(
                {
                    "error_code": "TARGET_FIELD_COLUMN_COUNT_INVALID",
                    "target_field": target,
                    "observed_count": len(columns),
                    "message_cn": "目标字段对应的源列数量不是1。",
                    "fix_suggestion": "删除重复映射或人工确认唯一源列。",
                }
            )
            continue
        confirmed_fields.append(
            {
                "target_field": target,
                "source_header": columns[0]["source_header"],
                "column_index": int(columns[0]["column_index"]),
                "column_letter": columns[0]["column_letter"],
                "confirmation_method": "PROFILE_RULE_CONFIRMATION",
                "confirmation_status": "CONFIRMED",
            }
        )

    used_columns = [item["column_index"] for item in confirmed_fields]
    if len(used_columns) != len(set(used_columns)):
        errors.append(
            {
                "error_code": "MANUAL_MAPPING_SOURCE_COLUMN_REUSED",
                "message_cn": "同一个源列不能映射到多个必需目标字段。",
                "fix_suggestion": "为每个目标字段选择不同的唯一源列。",
            }
        )

    status = "CONFIRMED" if not errors else "MANUAL_CONFIRMATION_REQUIRED"
    return {
        "status": status,
        "sheet_name": target_sheet,
        "header_row": sheet_report.get("header_row"),
        "requires_manual_confirmation": bool(errors),
        "user_override_count": len(overrides),
        "confirmed_fields": confirmed_fields,
        "errors": errors,
    }
