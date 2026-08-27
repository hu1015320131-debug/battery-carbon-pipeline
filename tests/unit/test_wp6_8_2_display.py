from carbon_excel_pipeline.ui.display_format import (
    format_activity_display,
    format_emission_display,
    format_full_precision,
    format_percentage_display,
)
from carbon_excel_pipeline.ui.reason_mapper import (
    UNKNOWN_REASON_LABEL,
    display_reason_code,
    display_route,
    display_severity,
    display_status,
)


def test_required_reason_codes_are_chinese():
    mapping = {
        "BUSINESS_UNIT_MISSING": "事业部信息缺失",
        "CHEMISTRY_MISSING": "化学体系信息缺失",
        "HISTORICAL_GHG_MISSING": "历史排放结果缺失",
        "HISTORICAL_GHG_VALUE_NON_NUMERIC": "历史排放结果不是有效数值",
        "HISTORICAL_GHG_VALUE_ZERO": "历史排放结果为 0",
        "MODEL_MISSING": "型号信息缺失",
        "PROJECT_MISSING": "项目信息缺失",
        "REPORTED_ACTIVITY_VALUE_ZERO": "活动数据为 0",
        "SUPPLIER_MISSING": "供应商信息缺失",
        "EF_VALUE_MISSING": "排放因子缺失",
        "MERGED_CELL_DATA_CONTEXT_DETECTED": "检测到合并单元格，程序已识别其上下文结构",
        "UNMAPPED_FIELDS_PRESENT": "检测到部分尚未识别的字段，原始字段已保留",
    }
    for code, label in mapping.items():
        assert display_reason_code(code) == label
        assert code not in display_reason_code(code)


def test_unknown_reason_code_uses_safe_fallback():
    assert display_reason_code("XXXX_XXXX_MISSING") == UNKNOWN_REASON_LABEL


def test_severity_and_capability_status_are_chinese():
    assert display_severity("Warning") == "需关注"
    assert display_severity("Blocking") == "阻断"
    assert display_status("PARTIALLY_CAPABLE") == "部分数据可处理"
    assert display_status("CAPABLE") == "可处理"
    assert display_status("INCAPABLE") == "当前数据无法完成处理"
    assert display_route("DIRECT_REPORTED_MASS") == "直接年度质量"
    assert display_route("SOURCE_EMBEDDED_FACTOR") == "文件内已提供"


def test_emission_display_uses_two_decimals_and_keeps_full_precision():
    raw = "0.200000"
    assert format_emission_display(raw) == "0.20 tCO2e"
    assert format_full_precision(raw) == raw
    assert format_activity_display("100.0000") == "100.00 kg/year"
    assert format_percentage_display("-5.5032", places=2) == "-5.50%"
    assert format_emission_display("-1340.887") == "-1,340.89 tCO2e"
