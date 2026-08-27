"""Unified UI labels for machine reason codes and statuses.

Backend payloads keep the original codes. Ordinary Streamlit pages must call
these helpers instead of rendering SCREAMING_SNAKE tokens.
"""

from __future__ import annotations

from typing import Any

UNKNOWN_REASON_LABEL = "存在需要检查的数据问题"

REASON_CODE_LABELS: dict[str, str] = {
    "BUSINESS_UNIT_MISSING": "事业部信息缺失",
    "CHEMISTRY_MISSING": "化学体系信息缺失",
    "SUPPLIER_MISSING": "供应商信息缺失",
    "PROJECT_MISSING": "项目信息缺失",
    "MODEL_MISSING": "型号信息缺失",
    "HISTORICAL_GHG_MISSING": "历史排放结果缺失",
    "HISTORICAL_GHG_VALUE_NON_NUMERIC": "历史排放结果不是有效数值",
    "HISTORICAL_GHG_VALUE_NEGATIVE": "历史排放结果为负数",
    "HISTORICAL_GHG_VALUE_ZERO": "历史排放结果为 0",
    "HISTORICAL_GHG_UNIT_MISSING": "历史排放结果缺少单位",
    "HISTORICAL_GHG_UNIT_UNSUPPORTED": "历史排放结果单位不受支持",
    "REPORTED_ACTIVITY_VALUE_ZERO": "活动数据为 0",
    "REPORTED_ACTIVITY_VALUE_MISSING": "活动数据缺失",
    "REPORTED_ACTIVITY_VALUE_NON_NUMERIC": "活动数据不是有效数值",
    "REPORTED_ACTIVITY_VALUE_NEGATIVE": "活动数据为负数",
    "REPORTED_ACTIVITY_UNIT_MISSING": "活动数据缺少单位",
    "REPORTED_ACTIVITY_UNIT_UNSUPPORTED": "活动数据单位不受支持",
    "REPORTED_ACTIVITY_FIELD_MISSING": "缺少活动数据字段",
    "EF_VALUE_MISSING": "排放因子缺失",
    "EF_VALUE_NON_NUMERIC": "排放因子不是有效数值",
    "EF_VALUE_NEGATIVE": "排放因子为负数",
    "EF_VALUE_ZERO": "排放因子为 0",
    "EF_UNIT_MISSING": "排放因子缺少单位",
    "EF_UNIT_UNSUPPORTED": "排放因子单位不受支持",
    "EF_SOURCE_MISSING": "排放因子来源缺失",
    "PCS_VALUE_MISSING": "采购数量缺失",
    "PCS_VALUE_NON_NUMERIC": "采购数量不是有效数值",
    "PCS_VALUE_NEGATIVE": "采购数量为负数",
    "PCS_VALUE_ZERO": "采购数量为 0",
    "PCS_FIELD_MISSING": "缺少采购数量字段",
    "UNIT_WEIGHT_VALUE_MISSING": "单件重量缺失",
    "UNIT_WEIGHT_VALUE_NON_NUMERIC": "单件重量不是有效数值",
    "UNIT_WEIGHT_VALUE_NEGATIVE": "单件重量为负数",
    "UNIT_WEIGHT_VALUE_ZERO": "单件重量为 0",
    "UNIT_WEIGHT_UNIT_MISSING": "单件重量缺少单位",
    "UNIT_WEIGHT_UNIT_UNSUPPORTED": "单件重量单位不受支持",
    "UNIT_WEIGHT_FIELD_MISSING": "缺少单件重量字段",
    "MERGED_CELL_DATA_CONTEXT_DETECTED": "检测到合并单元格，程序已识别其上下文结构",
    "UNMAPPED_FIELDS_PRESENT": "检测到部分尚未识别的字段，原始字段已保留",
    "DUPLICATE_SEMANTIC_MAPPING": "同一语义字段对应了多列，需要确认",
    "UNIT_MISSING_OR_UNKNOWN": "单位缺失或无法识别",
    "INSUFFICIENT_HEADER_EVIDENCE": "表头证据不足，已保留候选结果",
    "AMBIGUOUS_HEADER_CANDIDATES": "存在多个可能的表头位置",
    "WORKBOOK_UNRECOGNIZED": "未能识别当前工作簿结构",
    "CUSTOMER_UNMAPPED": "客户未完成映射",
    "CHEMISTRY_UNKNOWN": "化学体系未知",
    "SUPPLIER_UNMAPPED": "供应商未完成映射",
    "MISSING_REQUIRED_FIELD": "缺少必填字段",
    "PROJECT_MAPPING_PENDING": "项目映射待补充",
    "CELL_MODEL_UNKNOWN": "电芯型号未知",
    "LEGACY_CHEMISTRY_UNKNOWN": "化学体系未知",
    "LEGACY_CUSTOMER_UNMAPPED": "客户未完成映射",
    "NONE": "无",
    "FACTOR_NOT_AVAILABLE": "尚未找到适用排放因子",
    "BOUNDARY_POLICY_NOT_AVAILABLE": "尚未匹配核算范围",
    "PCS_OR_WEIGHT_INVALID": "采购数量或单件重量无效",
    "REQUESTED_ACTIVITY_PATH_NOT_SUPPORTED": "当前记录不支持所选活动数据路径",
}

SEVERITY_LABELS: dict[str, str] = {
    "Warning": "需关注",
    "WARNING": "需关注",
    "Blocking": "阻断",
    "BLOCKING": "阻断",
    "Info": "信息",
    "INFO": "信息",
    "ERROR": "错误",
    "PASS": "通过",
}

STATUS_LABELS: dict[str, str] = {
    "PASS": "通过",
    "PASS_EXACT": "一致",
    "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE": "通过，存在公式缓存精度差异",
    "DIFFERENCE_REQUIRES_REVIEW": "差异需复核",
    "HISTORICAL_NOT_AVAILABLE": "无历史结果可核对",
    "WARNING": "需关注",
    "ERROR": "错误",
    "BLOCKED": "无法继续核算",
    "PASS_WITH_WARNING": "通过（含需关注项）",
    "PARTIAL_RESULT": "部分结果可用",
    "OPEN": "待处理",
    "INDEPENDENT_CALCULATION_PASS": "独立复算通过",
    "INDEPENDENT_VALIDATION_FAIL": "独立复算未通过",
    "INDEPENDENT_VALIDATION_BLOCKED": "独立复算已阻断",
    "UNMATCHED": "未匹配",
    "MATCHED": "已匹配",
    "EXACT": "精确匹配",
    "FACTOR_NOT_AVAILABLE": "尚未找到适用排放因子",
    "NOT_AVAILABLE_FOR_SINGLE_INPUT_RUN": "单文件运行暂不可用",
    "NOT_RUN": "尚未运行",
    "ROUTED": "已完成路径选择",
    "UNKNOWN": "状态未知",
    "TRUE": "是",
    "FALSE": "否",
    "EXACT_MATCH": "精确匹配",
    "EXACT_LOCKED": "精确锁定",
    "CALCULATED_WITH_WARNING": "已核算（含限制说明）",
    "RECOGNIZED": "已识别",
    "RECOGNIZED_WITH_WARNING": "已识别，存在提示",
    "UNRECOGNIZED": "未识别",
    "UNMAPPED": "未识别",
    "MAPPED": "已识别",
    "CONFIRMED": "已确认",
    "DETECTED": "已检测",
    "CAPABLE": "可处理",
    "CAPABLE_WITH_WARNING": "可处理，存在提示",
    "PARTIALLY_CAPABLE": "部分数据可处理",
    "INCAPABLE": "当前数据无法完成处理",
    "G1A_UPSTREAM_REBUILD_RECONCILED": "上游重建对账通过",
    "DAY6_HISTORICAL_FACTOR_EXACT_MATCH_LOCKED": "历史模拟因子已精确匹配并锁定",
    "DAY7_CALCULATION_LINEAGE_RECONCILED": "碳核算与数据追溯对账通过",
    "G2_CLI_END_TO_END_PASS": "完整结果文件生成与验证通过",
}

ROUTE_LABELS: dict[str, str] = {
    "DIRECT_REPORTED_MASS": "直接年度质量",
    "PCS_WEIGHT_DERIVED": "数量 × 单件重量",
    "SOURCE_EMBEDDED_FACTOR": "文件内已提供",
    "HISTORICAL_SIMULATION_FACTOR": "历史模拟因子",
    "FACTOR_NOT_AVAILABLE": "尚未找到适用排放因子",
    "BOUNDARY_POLICY_NOT_AVAILABLE": "尚未匹配核算范围",
    "SYNTHETIC_DIRECT_MASS_BOUNDARY_V1": "已匹配",
    "PUBLIC_SYNTHETIC_DAY3_SCOPE_V1": "已匹配",
    "SOURCE_EMBEDDED_FACTOR_V1": "文件内已提供",
    "PUBLIC_SYNTHETIC_FACTOR_V1": "历史模拟因子",
    "SYNTHETIC_CELL_CATEGORY_FACTOR_V1": "2025 电芯类别历史模拟因子",
    "SYNTHETIC_CELL_CATEGORY_SCOPE_V1": "2025 电芯类别核算范围",
}


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _tokens(value: Any) -> list[str]:
    text = _text(value)
    if not text:
        return []
    for separator in (";", ",", "|"):
        text = text.replace(separator, " ")
    return [item for item in text.split() if item]


def display_reason_code(value: Any) -> str:
    tokens = _tokens(value)
    if not tokens:
        return "无"
    labels: list[str] = []
    for token in tokens:
        labels.append(REASON_CODE_LABELS.get(token, UNKNOWN_REASON_LABEL))
    unique: list[str] = []
    for label in labels:
        if label not in unique:
            unique.append(label)
    return "；".join(unique)


def display_severity(value: Any) -> str:
    text = _text(value)
    if not text:
        return ""
    return SEVERITY_LABELS.get(text, display_status(text))


def display_status(value: Any) -> str:
    text = _text(value)
    if not text:
        return ""
    if text in STATUS_LABELS:
        return STATUS_LABELS[text]
    if text in SEVERITY_LABELS:
        return SEVERITY_LABELS[text]
    if text in ROUTE_LABELS:
        return ROUTE_LABELS[text]
    if text in REASON_CODE_LABELS:
        return REASON_CODE_LABELS[text]
    if looks_like_machine_code(text):
        return UNKNOWN_REASON_LABEL
    return text


def display_historical_status(value: Any) -> str:
    text = _text(value)
    if text in {"PASS", "PASS_EXACT"}:
        return "一致"
    return display_status(value)


def display_route(value: Any) -> str:
    text = _text(value)
    if not text:
        return "未选择"
    return ROUTE_LABELS.get(text, display_status(text))


def looks_like_machine_code(value: Any) -> bool:
    text = _text(value)
    if not text or " " in text:
        return False
    if text.isupper() and "_" in text:
        return True
    return text.isupper() and text.isalpha() and len(text) > 3
