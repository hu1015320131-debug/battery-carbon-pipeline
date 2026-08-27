from decimal import Decimal
from pathlib import Path

from carbon_excel_pipeline.ui.business_view import (
    build_business_download_pack,
    comparison_ef_for_current_run,
    current_run_bound,
    factor_improvement_from_canonical,
    recognition_mapping_rows,
)

APP = Path(__file__).resolve().parents[2] / "app" / "streamlit_app.py"
OFFICIAL_2025_SHA = "SYNTHETIC_INPUT_SHA_PLACEHOLDER"


def _direct_mass_view() -> dict:
    return {
        "capability_summary": {
            "supported_activity_paths": ["DIRECT_REPORTED_MASS"],
            "direct_reported_mass_count": 2,
            "pcs_weight_derived_count": 0,
        },
        "semantic_mappings": [
            {
                "sheet_name": "（4.30）类别1外购原料辅料和服务",
                "field_mappings": [
                    {
                        "semantic_field": "Business_Unit",
                        "column_letter": "B",
                        "raw_header": "事业部",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": "Purchase_Type",
                        "column_letter": "C",
                        "raw_header": "采购类型",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": "Purchase_Category",
                        "column_letter": "D",
                        "raw_header": "公司外购原料和辅料名称",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": "Product_Description",
                        "column_letter": "E",
                        "raw_header": "物料描述",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": "Reported_Activity_Value",
                        "column_letter": "F",
                        "raw_header": "2024年购买的原料量（T/年）",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": "EF_Value",
                        "column_letter": "G",
                        "raw_header": "LCA排放因子",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": "EF_Source",
                        "column_letter": "H",
                        "raw_header": "排放因子来源",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": "Historical_GHG_Value",
                        "column_letter": "I",
                        "raw_header": "GHG排放量",
                        "mapping_status": "MAPPED",
                    },
                    {
                        "semantic_field": None,
                        "column_letter": "A",
                        "raw_header": "序号",
                        "mapping_status": "UNMAPPED",
                    },
                ],
            }
        ],
    }


def test_2024_recognition_mapping_does_not_default_to_serial_number():
    rows = recognition_mapping_rows(
        _direct_mass_view(),
        "（4.30）类别1外购原料辅料和服务",
        "DIRECT_REPORTED_MASS",
    )
    by_field = {item["业务字段"]: item for item in rows}
    assert by_field["事业部"]["原始表头"] == "事业部"
    assert by_field["采购类型"]["原始表头"] == "采购类型"
    assert by_field["物料类别"]["原始表头"] == "公司外购原料和辅料名称"
    assert by_field["物料描述"]["原始表头"] == "物料描述"
    assert "2024年购买的原料量（T/年）" in by_field["年度采购量"]["原始表头"]
    assert by_field["排放因子"]["原始表头"] == "LCA排放因子"
    assert by_field["排放因子来源"]["原始表头"] == "排放因子来源"
    assert by_field["历史排放量"]["原始表头"] == "GHG排放量"
    assert by_field["采购数量"]["识别列"] == "不适用"
    assert by_field["单件重量"]["识别列"] == "不适用"
    assert all("序号" not in item["原始表头"] for item in rows)
    assert all(item["识别列"] != "A列" for item in rows)


def test_unmapped_field_is_unrecognized_not_first_column():
    view = {
        "capability_summary": {"supported_activity_paths": ["DIRECT_REPORTED_MASS"]},
        "semantic_mappings": [
            {
                "sheet_name": "Sheet1",
                "field_mappings": [
                    {
                        "semantic_field": None,
                        "column_letter": "A",
                        "raw_header": "序号",
                        "mapping_status": "UNMAPPED",
                    }
                ],
            }
        ],
    }
    rows = recognition_mapping_rows(view, "Sheet1", "DIRECT_REPORTED_MASS")
    missing = next(item for item in rows if item["业务字段"] == "事业部")
    assert missing["识别列"] == "未识别"
    assert missing["原始表头"] != "序号"


def test_current_run_bound_requires_matching_id_and_sha():
    assert current_run_bound(
        run_id="RUN-A",
        input_sha256="abc",
        current_run_id="RUN-A",
        current_sha256="ABC",
    )
    assert not current_run_bound(
        run_id="RUN-A",
        input_sha256="abc",
        current_run_id="RUN-B",
        current_sha256="ABC",
    )
    assert not current_run_bound(
        run_id="RUN-A",
        input_sha256="aaa",
        current_run_id="RUN-A",
        current_sha256="BBB",
    )


def test_comparison_ef_is_gated_to_official_2025_sha():
    policy = comparison_ef_for_current_run(OFFICIAL_2025_SHA)
    assert policy is not None
    assert policy["comparison_ef"] == "2.500000"
    assert comparison_ef_for_current_run("DEADBEEF" * 8) is None


def test_factor_improvement_keeps_activity_constant():
    rows = [
        {
            "Activity_Data_kg": "100.0000",
            "EF_Value": "1.250000",
        }
    ]
    result = factor_improvement_from_canonical(rows, comparison_ef="2.500000")
    assert result is not None
    activity = Decimal(result["activity_kg"])
    simulated = Decimal(result["simulated_emission_kgco2e"])
    current = Decimal(result["current_emission_kgco2e"])
    assert activity == Decimal("100.0000")
    assert simulated == activity * Decimal("2.500000")
    assert current == activity * Decimal("1.250000")
    assert Decimal(result["reduction_kgco2e"]) == simulated - current
    assert current == Decimal("125.000000")


def test_business_download_pack_uses_chinese_names(tmp_path: Path):
    download = tmp_path / "08_download"
    download.mkdir()
    (download / "canonical_results.csv").write_text(
        "Record_ID,Product_Description,Business_Unit,Purchase_Category,Activity_Data_kg,Activity_Unit,Activity_Method,EF_Value,EF_Unit,EF_Source,Emission_kgCO2e,Calculation_QC,Warning_Codes\n"
        "REC-1,物料A,二部,电芯,1,kg/year,DIRECT_REPORTED_MASS,2.500000,kgCO2e/kg,文件内,2.500000,PASS,NONE\n",
        encoding="utf-8",
    )
    (download / "run_summary.json").write_text('{"Run_ID":"RUN-1","Input_File":"a.xlsx"}', encoding="utf-8")
    pack = build_business_download_pack(tmp_path)
    assert pack["cell_detail"]["display_name"] == "下载电芯数据明细"
    assert pack["carbon_result"]["display_name"] == "下载碳核算结果"
    assert pack["third_party"]["display_name"] == "下载第三方因子匹配输入表"
    assert pack["package"]["display_name"] == "下载完整结果包"
    assert pack["cell_detail"]["download_name"] == "电芯数据明细.csv"
    assert pack["package"]["download_name"] == "完整结果包.xlsx"
    detail = pack["cell_detail"]["data"].decode("utf-8-sig")
    assert "物料A" in detail
    assert "暂无数据" in pack["third_party"]["data"].decode("utf-8-sig")


def test_ordinary_ui_pages_and_isolation_guards():
    source = APP.read_text(encoding="utf-8")
    assert 'PAGES = [' in source
    assert '"数据导入与识别"' in source
    assert '"数据能力与核算范围"' in source
    assert '"数据质量与异常"' in source
    assert '"核算结果与分析"' in source
    assert '"结果下载"' in source
    assert "核算、验证与分析" not in source
    assert "清册衔接与下载" not in source
    assert "latest_wp6_4_run" not in source
    assert "latest_wp6_5_run" not in source
    assert "latest_wp6_6_run" not in source
    assert "latest_wp6_7_run" not in source
    assert "情景 A" not in source
    assert "A/B/C/D" not in source
    assert 'title=["排放量", "(tCO2e)"]' in source
    assert "labelAngle=0" in source
    assert "labelOverlap=False" in source
    assert "labelBound=False" in source
    assert "labels[0]" not in source
    assert "UNMAPPED_CHOICE" in source
    assert "st.metric(\"识别状态\"" not in source
    assert "st.metric(\"识别工作表\"" not in source
    assert '"WP6-3"' not in source
    assert '"WP6-4"' not in source
    assert "Frozen" not in source
    assert "Day3" not in source
    assert "Day7" not in source
    assert "LinkColumn" not in source
    assert "accept_multiple_files=True" in source
