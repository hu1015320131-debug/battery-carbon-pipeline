from pathlib import Path

APP = Path(__file__).resolve().parents[2] / "app" / "streamlit_app.py"


def test_ordinary_ui_has_no_local_path_mode_or_demo_title():
    source = APP.read_text(encoding="utf-8")
    assert "电芯数据碳核算程序" in source
    assert "电芯碳核算本地Demo" not in source
    assert "选择本地Excel" not in source
    assert "Excel文件路径" not in source
    assert "本机 Demo" not in source
    assert "本机Demo" not in source
    assert 'value=str(paths.raw_input)' not in source
    assert "labelAngle=0" in source
    assert "labelOverlap=False" in source
    assert "labelBound=False" in source
    assert "labelPadding=10" in source
    assert "format_emission_display" in source
    assert "display_reason_code" in source


def test_ordinary_ui_does_not_print_raw_reason_codes():
    source = APP.read_text(encoding="utf-8")
    forbidden = (
        "BUSINESS_UNIT_MISSING",
        "CHEMISTRY_MISSING",
        "EF_VALUE_MISSING",
        "PARTIALLY_CAPABLE",
        "MERGED_CELL_DATA_CONTEXT_DETECTED",
        "UNMAPPED_FIELDS_PRESENT",
    )
    for token in forbidden:
        assert f'"{token}"' not in source
        assert f"'{token}'" not in source
        assert f"{token}：" not in source
