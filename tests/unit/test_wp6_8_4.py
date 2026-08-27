from pathlib import Path

from openpyxl import Workbook

from carbon_excel_pipeline.io.header_detector import load_alias_config
from carbon_excel_pipeline.ui.business_view import (
    build_business_download_pack,
    current_run_bound,
    detected_unit_options,
    historical_validation_display_rows,
    record_detail_display,
)
from carbon_excel_pipeline.ui.reason_mapper import display_historical_status, display_status
from carbon_excel_pipeline.wp6_8_4.attribute_enrichment import (
    enrich_canonical_with_attributes,
)
from carbon_excel_pipeline.wp6_8_4.business_units import (
    detect_business_units,
    filter_by_business_unit,
    synthetic_supplier_subset,
    unit_label_from_sheet,
)
from carbon_excel_pipeline.wp6_8_4.file_roles import classify_mapped_fields, classify_workbook_role, reconcile_roles
from carbon_excel_pipeline.wp6_8_4.input_set import ROLE_ATTRIBUTE, ROLE_PRIMARY, input_set_sha256
from carbon_excel_pipeline.wp6_8_4.record_ids import assign_additional_record_ids, namespace_for_unit


ROOT = Path(__file__).resolve().parents[2]
APP = ROOT / "app" / "streamlit_app.py"
ALIAS = load_alias_config(ROOT / "config" / "import" / "field_aliases.json")


def _xlsx(path: Path, headers: list[str], rows: list[list[object]], sheet: str = "Sheet1") -> Path:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)
    return path


def test_file_roles_use_headers_not_filename(tmp_path: Path):
    primary = _xlsx(
        tmp_path / "属性补充.xlsx",
        ["事业部", "物料描述", "2024年购买的原料量（T/年）", "LCA排放因子"],
        [["合成一部", "电芯A", 1, 14.6]],
    )
    attribute = _xlsx(
        tmp_path / "主核算.xlsx",
        ["客户", "项目号", "电芯型号", "供应商简写", "电芯化学体系", "物料描述"],
        [["客户甲", "P1", "M1", "SYNA", "LFP", "电芯A"]],
    )
    primary_role = classify_workbook_role(primary, ALIAS)
    attribute_role = classify_workbook_role(attribute, ALIAS)
    assert primary_role["role"] == ROLE_PRIMARY
    assert attribute_role["role"] == ROLE_ATTRIBUTE
    reconciled = reconcile_roles([primary_role, attribute_role])
    assert [item["role"] for item in reconciled] == [ROLE_PRIMARY, ROLE_ATTRIBUTE]


def test_attribute_enrichment_never_overwrites_protected_fields():
    canonical = [
        {
            "Record_ID": "2025-DY2-SYNA-DX000001",
            "Product_Description": "电芯A",
            "Activity_Data_kg": "10",
            "EF_Value": "1.250000",
            "EF_Source": "历史模拟",
            "Emission_kgCO2e": "138.49386",
            "Chemistry": "",
        }
    ]
    attributes = [
        {
            "Product_Description": "电芯A",
            "Chemistry": "LFP",
            "Supplier": "供应商甲",
            "Activity_Data_kg": "999",
            "EF_Value": "1",
            "Emission_kgCO2e": "1",
            "Attribute_Source_File": "attr.xlsx",
            "Attribute_Source_Sheet": "二部2025",
            "Attribute_Source_Row": 2,
        }
    ]
    output, summary = enrich_canonical_with_attributes(canonical, attributes)
    row = output[0]
    assert row["Activity_Data_kg"] == "10"
    assert row["EF_Value"] == "1.250000"
    assert row["Emission_kgCO2e"] == "138.49386"
    assert row["Chemistry"] == "LFP"
    assert row["Supplier"] == "供应商甲"
    assert row["Attribute_Match_Method"] == "Product_Description"
    assert row["Attribute_Match_Status"] == "MATCHED"
    assert row["Attribute_Source_File"] == "attr.xlsx"
    assert summary["matched"] == 1


def test_attribute_match_is_exact_and_refuses_ambiguity():
    canonical = [
        {"Product_Description": "电芯A", "Activity_Data_kg": "1"},
        {"Product_Description": "电芯B", "Activity_Data_kg": "2"},
    ]
    attributes = [
        {"Product_Description": "电芯A", "Chemistry": "LFP"},
        {"Product_Description": "电芯A", "Chemistry": "LCO"},
    ]
    output, summary = enrich_canonical_with_attributes(canonical, attributes)
    assert output[0]["Attribute_Match_Status"] == "AMBIGUOUS"
    assert "Chemistry" not in output[0] or output[0].get("Chemistry") in {None, ""}
    assert output[1]["Attribute_Match_Status"] == "UNMATCHED"
    assert "部分补充属性无法可靠匹配" in summary["message"]


def test_input_set_sha_changes_when_any_file_changes():
    first = input_set_sha256([{"sha256": "AAA", "name": "a.xlsx"}, {"sha256": "BBB", "name": "b.xlsx"}])
    second = input_set_sha256([{"sha256": "AAA", "name": "a.xlsx"}, {"sha256": "CCC", "name": "b.xlsx"}])
    single = input_set_sha256([{"sha256": "AAA", "name": "a.xlsx"}])
    assert first != second
    assert first != single
    assert first == input_set_sha256([{"sha256": "BBB", "name": "b.xlsx"}, {"sha256": "AAA", "name": "a.xlsx"}])


def test_business_units_are_detected_and_filtered():
    rows = [
        {"Business_Unit": "合成二部", "Purchase_Category": "电芯.聚合物电芯", "Record_ID": "2024-DY2-SYNA-DX000001"},
        {"Business_Unit": "合成一部", "Purchase_Category": "电芯", "Record_ID": "2024-DY1-DX000001"},
        {"Business_Unit": "1", "Purchase_Category": "包装", "Record_ID": "NOISE"},
    ]
    detected = detect_business_units(rows)
    assert "合成二部" in detected
    assert "合成一部" in detected
    assert "1" not in detected
    six = filter_by_business_unit(rows, "合成二部")
    one = filter_by_business_unit(rows, "合成一部")
    assert len(six) == 1
    assert len(one) == 1
    assert len(filter_by_business_unit(rows, "全部")) == 3
    options = detected_unit_options(rows, detected)
    assert options[0] == "全部"
    assert "合成一部" in options


def test_additional_record_ids_do_not_reuse_synthetic_namespace():
    assert namespace_for_unit("合成一部") == "DY1"
    assert namespace_for_unit("合成二部") == "DY2"
    assigned = assign_additional_record_ids(
        [{"Source_Row": 20, "values": {}}, {"Source_Row": 9, "values": {}}],
        year=2024,
        namespace="DY1",
    )
    assert [item["Record_ID"] for item in assigned] == ["2024-DY1-DX000001", "2024-DY1-DX000002"]
    frozen = synthetic_supplier_subset(
        [
            {"Record_ID": "2024-DY2-SYNA-DX000001"},
            {"Record_ID": "2024-DY1-DX000001"},
            {"Record_ID": "2025-DY2-SYNA-DX000002"},
        ]
    )
    assert [row["Record_ID"] for row in frozen] == ["2024-DY2-SYNA-DX000001", "2025-DY2-SYNA-DX000002"]
    assert unit_label_from_sheet("一部") == "合成一部"
    assert unit_label_from_sheet("二部") == "合成二部"


def test_historical_table_and_statuses_are_chinese():
    rows = historical_validation_display_rows(
        [
            {
                "Record_ID": "2024-DY2-SYNA-DX000001",
                "Source_Row": "88",
                "Calculated_Emission_tCO2e": "1",
                "Historical_Emission_tCO2e": "1",
                "Difference_tCO2e": "0",
                "Validation_Status": "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE",
            }
        ]
    )
    text = str(rows)
    assert "记录编号" in rows[0]
    assert "原始数据行" in rows[0]
    assert "重新计算排放量（tCO2e）" in rows[0]
    assert "验证结果" in rows[0]
    assert "Calculated_Emission_tCO2e" not in text
    assert "Validation_Status" not in text
    assert "PASS_WITH_FORMULA_CACHE_PRECISION_DIFFERENCE" not in text
    assert rows[0]["验证结果"] == "通过，存在公式缓存精度差异"
    assert display_historical_status("PASS_EXACT") == "一致"
    assert display_status("BLOCKED") == "无法继续核算"
    assert display_status("PARTIAL_RESULT") == "部分结果可用"
    assert display_status("INDEPENDENT_CALCULATION_PASS") == "独立复算通过"


def test_record_detail_uses_business_labels():
    cards = record_detail_display(
        {
            "Record_ID": "2025-DY2-SYNA-DX000001",
            "Business_Unit": "合成二部",
            "Purchase_Category": "电芯",
            "Product_Description": "物料A",
            "Supplier": "SYNA",
            "Chemistry": "LFP",
            "Activity_Data_kg": "1",
            "Activity_Method": "PCS_WEIGHT_DERIVED",
            "EF_Value": "1.250000",
            "EF_Source": "历史模拟",
            "Emission_kgCO2e": "1.250000",
            "Calculation_QC": "PASS",
            "Governance_QC": "WARNING",
            "Source_File": "25年基础数据汇总 .xlsx",
            "Source_Sheet": "二部",
            "Source_Row": "12",
        }
    )
    labels = [item[0] for item in cards]
    assert "记录编号" in labels
    assert "事业部" in labels
    assert "供应商" in labels
    assert "化学体系" in labels
    assert "活动数据生成方式" in labels
    assert "原始文件" in labels
    assert "原始行号" in labels
    assert "Activity_Method" not in labels
    assert "Governance_QC" not in labels
    assert dict(cards)["记录编号"] == "2025-DY2-SYNA-DX000001"
    assert dict(cards)["供应商"] == "SYNA"
    assert "Source_Row" not in labels
    values = dict(cards)
    assert values["活动数据生成方式"] == "数量 × 单件重量"
    assert values["核算状态"] == "通过"
    assert values["数据质量状态"] == "需关注"


def test_current_run_bound_uses_combined_sha_when_present():
    assert current_run_bound(
        run_id="RUN-A",
        input_sha256="PRIMARY",
        current_run_id="RUN-A",
        current_sha256="PRIMARY",
        input_set_sha256="SET-1",
        current_set_sha256="SET-1",
    )
    assert not current_run_bound(
        run_id="RUN-A",
        input_sha256="PRIMARY",
        current_run_id="RUN-A",
        current_sha256="PRIMARY",
        input_set_sha256="SET-1",
        current_set_sha256="SET-2",
    )


def test_download_pack_filters_by_business_unit(tmp_path: Path):
    download = tmp_path / "08_download"
    download.mkdir()
    (download / "canonical_results.csv").write_text(
        "Record_ID,Product_Description,Business_Unit,Purchase_Category,Activity_Data_kg,Activity_Unit,Activity_Method,EF_Value,EF_Unit,EF_Source,Emission_kgCO2e,Calculation_QC,Warning_Codes\n"
        "A,物料A,合成一部,电芯,1,kg/year,DIRECT_REPORTED_MASS,1,kgCO2e/kg,文件内,1,PASS,NONE\n"
        "B,物料B,合成二部,电芯,2,kg/year,DIRECT_REPORTED_MASS,1,kgCO2e/kg,文件内,2,PASS,NONE\n",
        encoding="utf-8",
    )
    (download / "run_summary.json").write_text('{"Run_ID":"RUN-1","Input_File":"a.xlsx"}', encoding="utf-8")
    pack = build_business_download_pack(tmp_path, "合成一部")
    detail = pack["cell_detail"]["data"].decode("utf-8-sig")
    assert pack["cell_detail"]["display_name"] == "下载当前事业部电芯数据明细"
    assert "物料A" in detail
    assert "物料B" not in detail


def test_ordinary_pages_keep_session_and_avoid_record_links():
    source = APP.read_text(encoding="utf-8")
    assert "accept_multiple_files=True" in source
    assert "LinkColumn" not in source
    assert "record_id=" not in source
    assert "query_params" not in source
    assert "st.dataframe(historical.head(20)" not in source
    assert "historical_validation_display_rows" in source
    assert "selected_record_id" in source
    assert "_render_record_detail" in source
    assert "run_end_to_end_stage" in source
    detail_start = source.index("def _render_record_detail")
    detail_end = source.index("def _kpi_totals")
    assert "run_end_to_end_stage" not in source[detail_start:detail_end]


def test_classify_mapped_fields_prefers_activity_over_attributes():
    assert classify_mapped_fields({"Reported_Activity_Value", "Chemistry"}) == ROLE_PRIMARY
    assert classify_mapped_fields({"Chemistry", "Supplier", "Model"}) == ROLE_ATTRIBUTE
