"""Business-facing local Streamlit UI for the Excel carbon pipeline."""

from __future__ import annotations

import hashlib
import html
import json
from pathlib import Path
from typing import Any, Callable

import altair as alt
import pandas as pd
import streamlit as st

from carbon_excel_pipeline.ui.business_view import (
    activity_route_from_view,
    build_business_download_pack,
    comparison_ef_for_current_run,
    current_run_bound,
    detected_unit_options,
    factor_improvement_from_canonical,
    historical_validation_display_rows,
    recognition_mapping_rows,
    record_detail_display,
)
from carbon_excel_pipeline.io.header_detector import load_alias_config
from carbon_excel_pipeline.ui.day9_controller import (
    DEFAULT_LOCAL_CONFIG,
    REQUIRED_TARGETS,
    TARGET_FIELD_LABELS,
    business_stage_statuses,
    chemistry_display_order,
    default_mapping_overrides,
    inspect_stage,
    load_day9_paths,
    load_e2e_view,
    load_inspection_view,
    latest_wp6_8_run,
    load_processed_rows,
    load_quality_issue_rows,
    load_run_snapshot,
    load_wp6_3_view,
    load_wp6_8_view,
    path_config_for_display,
    run_calculation_and_export,
    run_cleaning_and_quality,
    run_end_to_end_stage,
    run_scope_stage,
    run_wp6_3_2024_stage,
    safe_error,
    save_uploaded_file,
    should_invalidate_run,
    source_identity,
    translate_issue_codes,
    translate_status,
    wp6_8_download_artifacts,
)
from carbon_excel_pipeline.wp6_8_4.business_units import filter_by_business_unit
from carbon_excel_pipeline.wp6_8_4.file_roles import classify_workbook_role, reconcile_roles
from carbon_excel_pipeline.wp6_8_4.input_set import (
    ROLE_ATTRIBUTE,
    ROLE_LEDGER,
    ROLE_PRIMARY,
    ROLE_UNKNOWN,
    input_set_sha256,
    primary_file,
)
from carbon_excel_pipeline.wp6_8_5.current_run import (
    clear_current_run_pointer,
    persist_current_run,
    restore_current_run,
)
from carbon_excel_pipeline.ui.display_format import (
    format_activity_display,
    format_emission_display,
    format_full_precision,
    format_percentage_display,
)
from carbon_excel_pipeline.ui.reason_mapper import (
    display_reason_code,
    display_route,
    display_severity,
    display_status,
)


st.set_page_config(
    page_title="电芯数据碳核算程序",
    page_icon="🧮",
    layout="wide",
)

PAGES = [
    "数据导入与识别",
    "数据能力与核算范围",
    "数据质量与异常",
    "核算结果与分析",
    "结果下载",
]
UNMAPPED_CHOICE = "未识别"


def _init_state() -> None:
    defaults = {
        "active_run_dir": None,
        "active_source_identity": None,
        "active_source_path": None,
        "last_result": None,
        "last_error": None,
        "last_operation": "尚未执行操作",
        "wp6_3_result": None,
        "current_input_fingerprint": None,
        "current_e2e_run_id": None,
        "business_page": PAGES[0],
        "Uploaded_File_Name": None,
        "Uploaded_File_SHA256": None,
        "Uploaded_File_Temp_Path": None,
        "Uploaded_Files": [],
        "Input_Set_SHA256": None,
        "selected_business_unit": "全部",
        "selected_issue_business_unit": "全部",
        "selected_record_id": None,
        "Current_Source_Path": None,
        "Current_Run_ID": None,
        "Current_Run_Root": None,
        "current_run_restore_warning": None,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


@st.dialog("处理状态", width="large")
def _operation_dialog() -> None:
    st.markdown("#### 当前状态")
    st.write(f"最近操作：{st.session_state.get('last_operation', '尚未执行操作')}")
    error = st.session_state.get("last_error")
    result = st.session_state.get("last_result")
    if error:
        detail = error.get("error", {})
        st.error("操作未能完成，请根据下面的说明处理后重试。")
        st.write(f"原因：{detail.get('message_cn', '请检查上传文件。')}")
        st.write(f"影响：{detail.get('impact', '当前操作未完成。')}")
        st.write(f"处理建议：{detail.get('fix_suggestion', '修正后重新运行。')}")
    elif result:
        st.success(f"完成状态：{display_status(result.get('status', 'PASS'))}")
        run_id = (
            result.get("run_id")
            or st.session_state.get("Current_Run_ID")
            or Path(result.get("run_directory", "")).name
        )
        if run_id:
            st.write(f"当前运行编号：{run_id}")
        route = result.get("route_decision") or {}
        if route:
            st.write(f"活动数据路径：{display_route(route.get('Activity_Route'))}")
            st.write(f"排放因子：{display_route(route.get('Factor_Route'))}")
            boundary_ready = route.get("Boundary_Ready")
            if boundary_ready is True or str(boundary_ready).upper() == "TRUE":
                st.write("边界状态：已匹配")
            else:
                st.write("边界状态：尚未匹配")
            if route.get("Emission_Ready") is False:
                st.info("当前文件可以完成活动数据整理，但尚未找到适用的排放因子，因此暂不能生成碳排放结果。")
            elif route.get("status") == "PARTIAL_RESULT":
                st.info("已完成已具备条件的处理，仍有核算范围或排放因子条件未满足。")
            else:
                st.write("当前可以继续查看核算结果。")
        file_name = st.session_state.get("Uploaded_File_Name")
        if file_name:
            st.write(f"当前文件：{file_name}")
    else:
        st.info("当前还没有处理结果。")
    with st.expander("高级说明"):
        st.write("原始文件不会被覆盖。历史模拟因子仅用于参考，不能作为生产决策依据。")
        st.write("因子用途：历史模拟；生产使用：否。")


def _clear_upload_session(*, keep_confirmation: bool = True, run_root: Path | None = None) -> None:
    for key in (
        "Uploaded_File_Name",
        "Uploaded_File_SHA256",
        "Uploaded_File_Temp_Path",
        "Uploaded_Files",
        "Input_Set_SHA256",
        "selected_business_unit",
        "selected_issue_business_unit",
        "selected_record_id",
        "Current_Source_Path",
        "Current_Run_ID",
        "Current_Run_Root",
        "current_run_restore_warning",
        "active_run_dir",
        "active_source_identity",
        "active_source_path",
        "current_input_fingerprint",
        "current_e2e_run_id",
        "last_result",
        "wp6_3_result",
    ):
        st.session_state[key] = None
    if not keep_confirmation:
        st.session_state["full_warning_confirmation"] = False
    if run_root is not None:
        clear_current_run_pointer(run_root)


def _execute(label: str, operation: Callable[[], dict[str, Any]]) -> dict[str, Any] | None:
    st.session_state["last_operation"] = label
    st.session_state["last_error"] = None
    with st.spinner(f"{label}正在执行，请勿关闭页面……"):
        try:
            result = operation()
        except Exception as error:
            st.session_state["last_error"] = safe_error(error)
            st.session_state["last_result"] = None
            _operation_dialog()
            return None
    st.session_state["last_result"] = result
    _operation_dialog()
    return result


def _set_active_run(result: dict[str, Any], source_path: Path) -> None:
    st.session_state["current_run_restore_warning"] = None
    run_value = result.get("run_directory")
    fingerprint = result.get("input_sha256") or source_identity(source_path)
    set_sha = str(result.get("input_set_sha256") or st.session_state.get("Input_Set_SHA256") or fingerprint).upper()
    resolved = str(source_path.resolve())
    if run_value:
        run_dir = str(Path(run_value).resolve())
        run_id = result.get("run_id") or Path(run_value).name
        st.session_state["active_run_dir"] = run_dir
        st.session_state["Current_Run_Root"] = run_dir
        st.session_state["current_e2e_run_id"] = run_id
        st.session_state["Current_Run_ID"] = run_id
        persist_current_run(Path(run_dir))
    st.session_state["active_source_identity"] = fingerprint
    st.session_state["current_input_fingerprint"] = fingerprint
    st.session_state["Uploaded_File_SHA256"] = str(fingerprint).upper()
    st.session_state["Input_Set_SHA256"] = set_sha
    st.session_state["active_source_path"] = resolved
    st.session_state["Current_Source_Path"] = resolved
    st.session_state["Uploaded_File_Temp_Path"] = resolved
    if not st.session_state.get("Uploaded_File_Name"):
        st.session_state["Uploaded_File_Name"] = source_path.name


def _current_source_path() -> Path | None:
    value = st.session_state.get("Current_Source_Path") or st.session_state.get(
        "Uploaded_File_Temp_Path"
    )
    if not value:
        return None
    path = Path(value)
    return path if path.is_file() else None


def _active_run_dir() -> Path | None:
    value = st.session_state.get("active_run_dir") or st.session_state.get("Current_Run_Root")
    if not value:
        return None
    path = Path(value)
    return path if path.is_dir() else None


def _session_sha() -> str:
    return str(
        st.session_state.get("Uploaded_File_SHA256")
        or st.session_state.get("current_input_fingerprint")
        or ""
    ).upper()


def _session_set_sha() -> str:
    return str(st.session_state.get("Input_Set_SHA256") or "").upper()


def _bound_e2e_view() -> dict[str, Any] | None:
    run_dir = _active_run_dir()
    if run_dir is None:
        return None
    view = load_e2e_view(run_dir)
    if not view:
        return None
    summary = view.get("e2e_summary") or {}
    if not current_run_bound(
        run_id=str(summary.get("Run_ID") or run_dir.name),
        input_sha256=str(summary.get("Input_SHA256") or ""),
        current_run_id=st.session_state.get("Current_Run_ID") or st.session_state.get("current_e2e_run_id"),
        current_sha256=_session_sha(),
        input_set_sha256=str(summary.get("Input_Set_SHA256") or ""),
        current_set_sha256=_session_set_sha(),
    ):
        return None
    return view


def _restore_completed_current_run(paths) -> None:
    if st.session_state.get("Current_Run_ID") or st.session_state.get("active_run_dir"):
        return
    restored = restore_current_run(paths.run_root)
    if not restored:
        return
    if restored.get("Restore_Status") == "LEGACY_RECORD_ID_SCHEMA":
        st.session_state["current_run_restore_warning"] = restored.get("Message")
        return
    for key, value in restored.items():
        st.session_state[key] = value
    run_dir = restored["Current_Run_Root"]
    primary_path = restored["Current_Source_Path"]
    primary_sha = restored["Uploaded_File_SHA256"]
    st.session_state["active_run_dir"] = run_dir
    st.session_state["active_source_path"] = primary_path
    st.session_state["active_source_identity"] = primary_sha
    st.session_state["current_input_fingerprint"] = primary_sha
    st.session_state["current_e2e_run_id"] = restored["Current_Run_ID"]


def _text_card(label: str, value: Any) -> None:
    text = "—" if value in (None, "") else str(value)
    st.markdown(
        (
            f'<div title="{html.escape(text)}" '
            'style="font-size:0.88rem;line-height:1.45;margin:0 0 0.75rem 0;">'
            f'<div style="color:#5c6770;font-size:0.75rem;">{html.escape(label)}</div>'
            '<div style="word-break:break-word;white-space:normal;overflow-wrap:anywhere;">'
            f"{html.escape(text)}</div></div>"
        ),
        unsafe_allow_html=True,
    )


def _bind_uploaded_files(uploaded_files, paths) -> None:
    st.session_state["current_run_restore_warning"] = None
    alias = load_alias_config(paths.project_root / "config" / "import" / "field_aliases.json")
    existing = {
        str(item.get("sha256")): dict(item)
        for item in (st.session_state.get("Uploaded_Files") or [])
        if item.get("sha256")
    }
    bound: list[dict[str, Any]] = []
    for uploaded in uploaded_files:
        content = uploaded.getvalue()
        digest = hashlib.sha256(content).hexdigest().upper()
        destination = save_uploaded_file(content, uploaded.name, paths)
        if digest in existing:
            item = existing[digest]
            item["name"] = uploaded.name
            item["path"] = str(destination.resolve())
            bound.append(item)
            continue
        classified = classify_workbook_role(destination, alias)
        classified["sha256"] = digest
        bound.append(classified)
    bound = reconcile_roles(bound)
    set_sha = input_set_sha256(bound)
    previous = st.session_state.get("Input_Set_SHA256") or st.session_state.get("Uploaded_File_SHA256")
    if should_invalidate_run(previous, set_sha):
        st.session_state["active_run_dir"] = None
        st.session_state["Current_Run_ID"] = None
        st.session_state["Current_Run_Root"] = None
        st.session_state["current_e2e_run_id"] = None
        st.session_state["last_result"] = None
        st.session_state["wp6_3_result"] = None
        st.session_state["selected_record_id"] = None
        st.info("已更换文件，先前结果不再展示，请重新检查或核算。")
    primary = primary_file(bound) or bound[0]
    st.session_state["Uploaded_Files"] = bound
    st.session_state["Input_Set_SHA256"] = set_sha
    st.session_state["Uploaded_File_Name"] = primary.get("name")
    st.session_state["Uploaded_File_SHA256"] = str(primary.get("sha256") or "").upper()
    st.session_state["Uploaded_File_Temp_Path"] = primary.get("path")
    st.session_state["Current_Source_Path"] = primary.get("path")
    st.session_state["active_source_identity"] = primary.get("sha256")
    st.session_state["current_input_fingerprint"] = primary.get("sha256")
    st.session_state["active_source_path"] = primary.get("path")


def _selected_source(paths) -> Path | None:
    st.markdown("#### 上传数据文件")
    uploaded = st.file_uploader(
        "上传 Excel（可一次选择多个）",
        type=["xlsx"],
        key="raw_upload",
        accept_multiple_files=True,
        help="可同时上传主核算数据、属性补充数据和历史清册/因子参考。",
    )
    if uploaded:
        _bind_uploaded_files(uploaded, paths)
    files = list(st.session_state.get("Uploaded_Files") or [])
    explicit_source = _current_source_path()
    if files and explicit_source is not None:
        listed_paths = {
            str(Path(str(item.get("path"))).resolve())
            for item in files
            if item.get("path")
        }
        if str(explicit_source.resolve()) not in listed_paths:
            files = [
                {
                    "name": st.session_state.get("Uploaded_File_Name") or explicit_source.name,
                    "sha256": st.session_state.get("Uploaded_File_SHA256"),
                    "path": str(explicit_source.resolve()),
                    "role": ROLE_PRIMARY,
                }
            ]
            st.session_state["Uploaded_Files"] = files
    if not files:
        source = _current_source_path()
        file_name = st.session_state.get("Uploaded_File_Name")
        if source is not None and file_name:
            files = [
                {
                    "name": file_name,
                    "sha256": st.session_state.get("Uploaded_File_SHA256"),
                    "path": str(source),
                    "role": ROLE_PRIMARY,
                }
            ]
            st.session_state["Uploaded_Files"] = files
    if files:
        markers = "①②③④⑤⑥⑦⑧⑨⑩"
        st.success(f"已上传 {len(files)} 个文件")
        updated: list[dict[str, Any]] = []
        for index, item in enumerate(files):
            marker = markers[index] if index < len(markers) else str(index + 1)
            st.write(f"{marker} {item.get('name')}")
            choices = [ROLE_PRIMARY, ROLE_ATTRIBUTE, ROLE_LEDGER, ROLE_UNKNOWN]
            current_role = item.get("role") if item.get("role") in choices else ROLE_UNKNOWN
            role = st.selectbox(
                "识别用途",
                choices,
                index=choices.index(current_role),
                key=f"file_role_{item.get('sha256') or index}",
            )
            item = dict(item)
            item["role"] = role
            updated.append(item)
        st.session_state["Uploaded_Files"] = updated
        primary = primary_file(updated)
        if primary and primary.get("path"):
            st.session_state["Uploaded_File_Name"] = primary.get("name")
            st.session_state["Uploaded_File_SHA256"] = str(primary.get("sha256") or "").upper()
            st.session_state["Uploaded_File_Temp_Path"] = primary.get("path")
            st.session_state["Current_Source_Path"] = primary.get("path")
            st.session_state["Input_Set_SHA256"] = input_set_sha256(updated)
        if st.button("清除当前文件", key="clear_uploaded_source"):
            _clear_upload_session(run_root=paths.run_root)
            st.rerun()
        source = _current_source_path()
        return source
    st.info("请上传 Excel 文件后开始处理。")
    return None


def _extra_source_paths() -> list[Path]:
    files = st.session_state.get("Uploaded_Files") or []
    extras: list[Path] = []
    for item in files:
        if item.get("role") == ROLE_PRIMARY or not item.get("path"):
            continue
        path = Path(str(item["path"]))
        if path.is_file():
            extras.append(path)
    return extras


def _file_role_map() -> dict[str, str]:
    return {
        str(item.get("name")): str(item.get("role"))
        for item in (st.session_state.get("Uploaded_Files") or [])
        if item.get("name") and item.get("role")
    }


def _preview_frame(preview: dict[str, Any]) -> pd.DataFrame:
    rows = [["" if value is None else str(value) for value in row["values"]] for row in preview["rows"]]
    if not rows:
        return pd.DataFrame()
    headers = [value or f"第{index}列" for index, value in enumerate(rows[0], start=1)]
    return pd.DataFrame(rows[1:], columns=headers)


def _render_collection_page(paths) -> None:
    st.subheader("数据收集、上传与一键核算")
    st.write(
        "上传电芯采购 Excel 后，可以先检查数据结构，也可以直接开始核算。"
        "程序会识别文件结构和数据能力；已配置核算范围和排放因子的文件可以完成碳排放计算，"
        "否则会说明还缺少哪些条件。"
    )
    source = _selected_source(paths)

    accepted = st.checkbox(
        "我了解核算结果仅供内部参考，不作为生产决策依据",
        key="full_warning_confirmation",
    )
    left, right = st.columns(2)
    with left:
        if st.button("检查并预览数据", type="primary", disabled=source is None, width="stretch"):
            result = _execute("数据接收与结构检查", lambda: inspect_stage(source, paths))
            if result:
                _set_active_run(result, source)
    with right:
        if st.button(
            "开始处理 / 开始核算",
            disabled=source is None or not accepted,
            width="stretch",
        ):
            existing = Path(st.session_state["active_run_dir"]) if st.session_state.get("active_run_dir") else None
            extras = _extra_source_paths()
            roles = _file_role_map()
            result = _execute(
                "端到端一键核算",
                lambda: run_end_to_end_stage(
                    source,
                    paths,
                    existing,
                    extra_source_paths=extras,
                    file_roles=roles,
                ),
            )
            if result:
                _set_active_run(result, source)

    active = _active_run_dir()
    if active is None:
        return
    view = load_inspection_view(active)
    st.markdown("#### 当前数据预览")
    recognition = view.get("recognition_summary") or {}
    if recognition:
        best_sheet = recognition.get("best_candidate_sheet") or "未找到"
        best_header = recognition.get("best_candidate_header_row") or "未确认"
        info_a, info_b, info_c = st.columns(3)
        with info_a:
            st.metric("工作表数量", recognition.get("sheet_count", 0))
        with info_b:
            _text_card("识别状态", display_status(recognition.get("recognition_status", "UNRECOGNIZED")))
        with info_c:
            _text_card("识别工作表", best_sheet)
        _text_card("表头位置", f"第 {best_header} 行")
        file_name = st.session_state.get("Uploaded_File_Name")
        if file_name:
            _text_card("文件名", file_name)
    inventory = pd.DataFrame(view["inventory"]).rename(
        columns={
            "sheet_name": "工作表",
            "physical_row_count": "总行数",
            "column_count": "列数",
            "header_detected": "已识别表头",
            "header_row": "表头行",
            "data_row_count": "数据记录数",
            "formula_count": "公式数量",
            "merged_cell_count": "合并单元格数量",
            "recognition_status": "识别状态",
            "recognized_field_count": "识别字段数",
        }
    )
    visible_columns = [
        "工作表",
        "总行数",
        "列数",
        "已识别表头",
        "表头行",
        "数据记录数",
        "公式数量",
        "合并单元格数量",
        "识别状态",
        "识别字段数",
    ]
    if "识别状态" in inventory.columns:
        inventory["识别状态"] = inventory["识别状态"].map(display_status)
    st.dataframe(
        inventory[[item for item in visible_columns if item in inventory.columns]],
        hide_index=True,
        width="stretch",
    )
    best_mapping = next(
        (
            item
            for item in view.get("semantic_mappings", [])
            if item.get("sheet_name") == recognition.get("best_candidate_sheet")
        ),
        None,
    )
    if best_mapping:
        st.markdown("#### 最佳候选字段识别结果")
        mapping_rows = best_mapping.get("field_mappings", [])
        mapping_frame = pd.DataFrame(mapping_rows).rename(
            columns={
                "raw_header": "原始字段",
                "semantic_field": "识别语义",
                "detected_unit": "单位",
                "mapping_status": "状态",
                "warning_code": "提示",
            }
        )
        if "提示" in mapping_frame.columns:
            mapping_frame["提示"] = mapping_frame["提示"].map(display_reason_code)
        if "状态" in mapping_frame.columns:
            mapping_frame["状态"] = mapping_frame["状态"].map(display_status)
        st.dataframe(
            mapping_frame[["原始字段", "识别语义", "单位", "状态", "提示"]],
            hide_index=True,
            width="stretch",
        )
        unmapped_count = sum(1 for item in mapping_rows if item.get("mapping_status") == "UNMAPPED")
        st.caption(f"未识别字段：{unmapped_count} 个")
        sheet_summary = next(
            (
                item
                for item in recognition.get("sheets", [])
                if item.get("sheet_name") == recognition.get("best_candidate_sheet")
            ),
            {},
        )
        for warning in sheet_summary.get("warnings", []):
            st.warning(display_reason_code(warning.get("code")))
    capability = view.get("capability_summary") or {}
    if capability:
        st.markdown("#### 数据能力识别")
        st.caption("这里只判断数据能否进入核算，不会在本步生成最终排放结果。")
        cap_a, cap_b, cap_c, cap_d = st.columns(4)
        cap_a.metric("活动数据就绪", capability.get("activity_ready_count", 0))
        cap_b.metric("排放因子就绪", capability.get("factor_ready_count", 0))
        cap_c.metric("可计算排放", capability.get("emission_ready_count", 0))
        cap_d.metric("可核对历史结果", capability.get("historical_validation_ready_count", 0))
        path_a, path_b = st.columns(2)
        path_a.metric(
            "数量 × 单重路径",
            capability.get("pcs_weight_derived_count", 0),
            f"覆盖率 {capability.get('pcs_weight_derived_coverage', 0):.2%}",
        )
        path_b.metric(
            "直接年度质量路径",
            capability.get("direct_reported_mass_count", 0),
            f"覆盖率 {capability.get('direct_reported_mass_coverage', 0):.2%}",
        )
        st.write(
            f"数据能力：{display_status(capability.get('status', 'UNKNOWN'))}；"
            f"数据记录：{capability.get('total_records', 0)} 条。"
        )
        reason_rows = [
            {"类型": display_severity("Warning"), "问题": display_reason_code(code), "记录数": count}
            for code, count in capability.get("warning_code_counts", {}).items()
        ] + [
            {"类型": display_severity("Blocking"), "问题": display_reason_code(code), "记录数": count}
            for code, count in capability.get("blocking_code_counts", {}).items()
        ]
        if reason_rows:
            st.dataframe(pd.DataFrame(reason_rows), hide_index=True, width="stretch")
        if (
            capability.get("direct_reported_mass_count", 0) > 0
            and capability.get("historical_validation_ready_count", 0) > 0
        ):
            with st.expander("高级验证 / 历史复现工具（普通一键流程已自动核算，不必再点）"):
                st.caption("正式一键核算已自动选择直接年度质量路径。此按钮仅保留给需要单独复跑历史复现的高级验证。")
                if st.button("执行历史复现", key="run_historical_reproduction"):
                    result = _execute(
                        "历史复现核算",
                        lambda: run_wp6_3_2024_stage(active, paths),
                    )
                    if result:
                        st.session_state["wp6_3_result"] = result
                wp63_result = st.session_state.get("wp6_3_result") or {}
                output_dir = wp63_result.get("output_directory")
                if output_dir and Path(output_dir).is_dir():
                    wp63 = load_wp6_3_view(Path(output_dir))
                    summary = wp63["summary"]
                    scope_a, scope_b, scope_c = st.columns(3)
                    scope_a.metric("二部", summary["boundary"]["business_unit_count"])
                    scope_b.metric("电芯", summary["boundary"]["cell_count"])
                    scope_c.metric("试点范围", summary["boundary"]["marker_count"])
                    total_a, total_b, total_c = st.columns(3)
                    total_a.metric("活动数据 kg/year", summary["totals"]["activity_kg"])
                    total_b.metric("重新计算 tCO2e", summary["totals"]["calculated_emission_tco2e"])
                    total_c.metric("与历史差异 tCO2e", summary["totals"]["difference_tco2e"])
    selected_sheet = st.selectbox(
        "选择预览工作表",
        [item["sheet_name"] for item in view["inventory"]],
        key="preview_sheet",
    )
    preview = next(item for item in view["previews"] if item["sheet_name"] == selected_sheet)
    st.dataframe(_preview_frame(preview), hide_index=True, width="stretch")


def _record_id_explanation() -> None:
    st.markdown("#### 记录编号是什么？")
    st.info(
        "记录编号是每条电芯采购记录在整个核算流程中的唯一身份证。"
        "同一条源记录重复运行仍得到同一个编号。"
        "编号格式为“年份-事业部-供应商-物料代码+6位流水号”；供应商未知时使用 UNK。"
        "它连接清洗结果、活动数据、排放因子、匹配、核算和追溯证据。"
    )


def _render_mapping_page(paths) -> None:
    st.subheader("字段映射与核算范围")
    active = _active_run_dir()
    if active is None:
        st.info("请先在「数据导入与识别」页面检查原始 Excel。")
        _record_id_explanation()
        return
    view = load_inspection_view(active)
    recognition = view.get("recognition_summary") or {}
    semantic_sheets = [item.get("sheet_name") for item in view.get("semantic_mappings", []) if item.get("sheet_name")]
    fallback_sheets = [item["sheet_name"] for item in view["mappings"] if item.get("detected")]
    sheets = semantic_sheets or fallback_sheets
    if not sheets:
        st.warning("当前工作簿没有可进入后续流程的工作表。未识别字段将显示为「未识别」，不会默认使用第一列。")
        _record_id_explanation()
        return
    best = recognition.get("best_candidate_sheet")
    default_index = sheets.index(best) if best in sheets else 0
    target_sheet = st.selectbox("处理工作表", sheets, index=default_index)
    route = activity_route_from_view(view)
    if route:
        st.caption(f"当前识别路径：{display_route(route)}")
    rows = recognition_mapping_rows(view, target_sheet, route)
    st.markdown("#### 识别字段对照")
    st.caption("下表直接读取识别结果。未识别字段显示「未识别」，不会默认映射到第一列。")
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    if any(item["状态"] == "未识别" for item in rows):
        st.info("部分业务字段尚未识别。直接质量路径不需要采购数量或单件重量时，会显示「不适用」。")
    _record_id_explanation()

    with st.expander("高级：手工确认映射并筛选核算范围"):
        auto = default_mapping_overrides(view, target_sheet) if target_sheet in view.get("columns", {}) else {}
        source_columns = view.get("columns", {}).get(target_sheet, [])
        labels = [
            f'{item["column_index"]}|{item["column_letter"]}|{item["source_header"]}'
            for item in source_columns
        ]
        choices = [UNMAPPED_CHOICE, *labels]
        mapping_overrides: dict[str, dict[str, Any]] = {}
        st.caption("未识别字段保持「未识别」，不会自动选中序号或其他第一列。")
        for target in REQUIRED_TARGETS:
            default = auto.get(target)
            default_label = (
                f'{default["column_index"]}|{default["column_letter"]}|{default["source_header"]}'
                if default
                else UNMAPPED_CHOICE
            )
            choice = st.selectbox(
                TARGET_FIELD_LABELS[target],
                choices,
                index=choices.index(default_label) if default_label in choices else 0,
                key=f"mapping_{target}",
            )
            if choice == UNMAPPED_CHOICE:
                continue
            index_text, letter, header = choice.split("|", 2)
            mapping_overrides[target] = {
                "column_index": int(index_text),
                "column_letter": letter,
                "source_header": header,
            }
        category = st.text_input(
            "进入核算范围的采购分类",
            value="电芯.聚合物电芯.聚合物电芯",
            key="target_category",
        )
        markers_text = st.text_input("供应商识别标记（逗号分隔）", value="SYNA", key="supplier_markers")
        markers = [item.strip() for item in markers_text.split(",") if item.strip()]
        confirmed = st.checkbox("我确认工作表、字段映射和核算范围", key="mapping_confirmed")
        snapshot = load_run_snapshot(active, paths)
        completed = "day3" in snapshot["technical_status"]
        if completed:
            st.caption("如需修改映射或范围，请回到第一页重新检查原始 Excel 并创建新的运行编号。")
        if st.button("确认映射并筛选核算范围", disabled=not confirmed or not markers or completed):
            _execute(
                "字段映射确认与范围筛选",
                lambda: run_scope_stage(
                    active,
                    paths,
                    target_sheet=target_sheet,
                    target_purchase_category=category,
                    supplier_markers=markers,
                    mapping_overrides=mapping_overrides,
                ),
            )


def _render_cleaning_page(paths) -> None:
    st.subheader("数据质量与异常")
    st.write("本页只展示当前上传文件的质量、验证和异常信息。没有当前结果时显示「暂无数据」。")
    e2e = _bound_e2e_view()
    if e2e is None:
        st.info("暂无数据")
        st.caption("请先在「数据导入与识别」完成当前文件的检查或核算。")
        with st.expander("高级：清洗规则说明"):
            _render_quality_rules()
        return

    summary = e2e.get("e2e_summary") or {}
    canonical = e2e.get("canonical") or []
    _text_card("当前文件", summary.get("Input_File") or st.session_state.get("Uploaded_File_Name"))
    _text_card("识别工作表", summary.get("Selected_Sheet"))
    _text_card("处理路径", display_route(summary.get("Activity_Route")))
    _text_card("排放因子来源", display_route(summary.get("Factor_Route")))

    calc_pass = sum(1 for row in canonical if str(row.get("Calculation_QC", "")).upper() == "PASS")
    gov_pass = sum(1 for row in canonical if str(row.get("Governance_QC", "")).upper() == "PASS")
    gov_warn = sum(1 for row in canonical if str(row.get("Governance_QC", "")).upper() == "WARNING")
    boundary_ready = sum(1 for row in canonical if str(row.get("Boundary_Ready", "")).upper() in {"TRUE", "1"})
    cols = st.columns(4)
    cols[0].metric("核算质量通过", f"{calc_pass} / {len(canonical)}")
    cols[1].metric("治理完整", f"{gov_pass} / {len(canonical)}")
    cols[2].metric("需关注", gov_warn)
    cols[3].metric("边界可判定", f"{boundary_ready} / {len(canonical)}")

    st.markdown("#### 独立计算验证")
    st.metric("验证状态", display_status(summary.get("Independent_Validation_Status", "NOT_RUN")))
    download = e2e.get("download_dir")
    if download:
        score_path = Path(download) / "quality_scorecards.json"
        issue_path = Path(download) / "data_quality_issue_register.csv"
        validation_path = Path(download) / "validation_results.csv"
        historical_path = Path(download) / "historical_validation.csv"
        if score_path.is_file():
            scorecards = json.loads(score_path.read_text(encoding="utf-8"))
            if isinstance(scorecards, list):
                for card in scorecards:
                    st.caption(f"质量分层 · {card.get('Year', '当前')}")
                    metrics = {item["Metric"]: item for item in card.get("metrics", [])}
                    metric_labels = {
                        "Calculation_Readiness": "核算就绪",
                        "Governance_Field_Completeness": "治理字段完整",
                        "Boundary_Readiness": "边界就绪",
                        "Traceability": "可追溯性",
                    }
                    metric_cols = st.columns(4)
                    for column, name in zip(metric_cols, metric_labels):
                        metric = metrics.get(name)
                        if not metric:
                            column.metric(metric_labels[name], "暂无数据")
                            continue
                        column.metric(
                            metric_labels[name],
                            f"{metric['Numerator']} / {metric['Denominator']}",
                            help=f"完整比率：{format_full_precision(metric.get('Rate'))}",
                        )
        if validation_path.is_file():
            validation_frame = pd.read_csv(validation_path, dtype=str).fillna("")
            if not validation_frame.empty:
                show_cols = [
                    item
                    for item in (
                        "Record_ID",
                        "Activity_Validation_Status",
                        "EF_Validation_Status",
                        "Emission_Validation_Status",
                        "Overall_Validation_Status",
                    )
                    if item in validation_frame.columns
                ]
                renamed = validation_frame[show_cols].rename(
                    columns={
                        "Record_ID": "记录编号",
                        "Activity_Validation_Status": "活动数据验证",
                        "EF_Validation_Status": "排放因子验证",
                        "Emission_Validation_Status": "排放量验证",
                        "Overall_Validation_Status": "总体验证",
                    }
                )
                for column in renamed.columns:
                    if column != "记录编号":
                        renamed[column] = renamed[column].map(display_status)
                with st.expander("查看逐条验证结果"):
                    st.dataframe(renamed, hide_index=True, width="stretch")
        if historical_path.is_file():
            historical = pd.read_csv(historical_path, dtype=str).fillna("")
            st.markdown("#### 历史结果验证")
            if historical.empty:
                st.info("暂无数据")
            else:
                st.caption("仅核对本文件已有历史排放结果，不引入其他年度运行。")
                st.dataframe(
                    pd.DataFrame(historical_validation_display_rows(historical.to_dict(orient="records"))),
                    hide_index=True,
                    width="stretch",
                )
        if issue_path.is_file():
            issues = pd.read_csv(issue_path, dtype=str).fillna("")
            st.markdown("#### 异常与需关注项")
            if issues.empty:
                st.success("当前记录没有登记问题。")
            else:
                issues = issues.copy()
                if "Business_Unit" in issues.columns:
                    issue_units = ["全部", *[item for item in issues["Business_Unit"].drop_duplicates().tolist() if item]]
                    current_issue_unit = st.session_state.get("selected_issue_business_unit") or "全部"
                    if current_issue_unit not in issue_units:
                        current_issue_unit = "全部"
                    selected_issue_unit = st.selectbox(
                        "按事业部筛选异常",
                        issue_units,
                        index=issue_units.index(current_issue_unit),
                        key="selected_issue_business_unit",
                    )
                    if selected_issue_unit != "全部":
                        issues = issues[issues["Business_Unit"] == selected_issue_unit]
                if "Issue_Code" in issues.columns:
                    issues["问题"] = issues["Issue_Code"].map(display_reason_code)
                if "Severity" in issues.columns:
                    issues["程度"] = issues["Severity"].map(display_severity)
                display_frame = issues.rename(columns={"Business_Unit": "事业部", "Record_ID": "记录编号", "Priority": "优先级", "Description": "说明"})
                keep = [item for item in ("事业部", "记录编号", "问题", "程度", "优先级", "说明") if item in display_frame.columns]
                st.dataframe(display_frame[keep], hide_index=True, width="stretch")

    run_dir = _active_run_dir()
    if run_dir is not None:
        with st.expander("高级：运行进度与清洗规则"):
            snapshot = load_run_snapshot(run_dir, paths)
            st.dataframe(pd.DataFrame(business_stage_statuses(snapshot)), hide_index=True, width="stretch")
            quality = snapshot["upstream_quality"]
            qcols = st.columns(3)
            qcols[0].metric("通过", quality["PASS"])
            qcols[1].metric("警告", quality["WARNING"])
            qcols[2].metric("错误", quality["ERROR"])
            _render_quality_rules()
            issues = load_quality_issue_rows(run_dir)
            if issues:
                st.markdown("清洗质检问题")
                st.dataframe(pd.DataFrame(issues), hide_index=True, width="stretch")
            processed = load_processed_rows(run_dir)
            if processed:
                st.markdown("清洗后数据")
                st.dataframe(pd.DataFrame(processed), hide_index=True, width="stretch")
            if st.button("开始清洗和质检", key="manual_cleaning"):
                technical = snapshot["technical_status"]
                if technical.get("day3") == "PASS" and "day5" not in technical:
                    _execute("清洗、标准化与质量检查", lambda: run_cleaning_and_quality(run_dir, paths))
                else:
                    st.info("当前一键流程已包含清洗，或尚不具备分步清洗条件。")


def _render_quality_rules() -> None:
    rules = pd.DataFrame(
        [
            {"检查内容": "文件和表头", "可以通过": "标准.xlsx、文件可读取、必需字段唯一", "不能通过及原因": "损坏、加密、缺列或一列映射多个字段，无法确定数据含义"},
            {"检查内容": "采购数量", "可以通过": "严格大于0的整数", "不能通过及原因": "空白、文字、0、负数或小数，无法代表有效采购件数"},
            {"检查内容": "单件重量和采购量", "可以通过": "严格大于0的十进制数", "不能通过及原因": "空白、文字、0、负数、无穷值，无法计算活动量"},
            {"检查内容": "单位", "可以通过": "配置允许的PCS、g/PCS、g/year等精确单位", "不能通过及原因": "未知单位、大小写错误或多余空格，可能造成错误换算"},
            {"检查内容": "核算范围", "可以通过": "目标采购分类且物料信息包含供应商标记", "不能通过及原因": "不属于试点范围，记录会进入排除审计而非直接删除"},
            {"检查内容": "记录编号", "可以通过": "非空、唯一且可回溯源行", "不能通过及原因": "重复、缺失或无法追溯，不能安全连接因子和核算结果"},
            {"检查内容": "主数据映射", "可以通过": "供应商、项目等已映射；部分缺失可作为警告", "不能通过及原因": "关键主数据冲突会阻断；非关键缺失保留为警告"},
        ]
    )
    st.dataframe(rules, hide_index=True, width="stretch")


def _summary_bar_chart(
    frame: pd.DataFrame,
    *,
    category: str,
    category_order: list[str] | str,
) -> alt.Chart:
    return (
        alt.Chart(frame)
        .mark_bar(color="#5B9BD5", cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
        .encode(
            x=alt.X(
                f"{category}:N",
                sort=category_order,
                axis=alt.Axis(
                    title=category,
                    labelAngle=0,
                    labelOverlap=False,
                    labelBound=False,
                    labelPadding=10,
                    titleAngle=0,
                    labelLimit=200,
                ),
            ),
            y=alt.Y(
                "排放量:Q",
                axis=alt.Axis(
                    title=["排放量", "(tCO2e)"],
                    labelAngle=0,
                    titleAngle=0,
                    titleAlign="left",
                    titleAnchor="end",
                    titleX=0,
                    titleY=-18,
                ),
            ),
            tooltip=[
                alt.Tooltip(f"{category}:N", title=category),
                alt.Tooltip("排放量:Q", title="排放量（tCO2e）", format=",.6f"),
                alt.Tooltip("完整值:N", title="完整值"),
            ],
        )
        .properties(height=310)
        .configure_axis(labelFontSize=12, titleFontSize=12)
    )


def _chart_frame(canonical: list[dict[str, Any]], *, category_field: str, category_label: str) -> pd.DataFrame:
    rows = []
    for item in canonical:
        label = str(item.get(category_field) or "").strip() or "未填写"
        emission = item.get("Emission_tCO2e") or item.get("Emission_kgCO2e")
        try:
            value = float(str(emission))
        except (TypeError, ValueError):
            continue
        if "Emission_tCO2e" not in item and item.get("Emission_kgCO2e"):
            value = value / 1000.0
        rows.append(
            {
                category_label: label,
                "排放量": value,
                "完整值": format_full_precision(item.get("Emission_tCO2e") or item.get("Emission_kgCO2e")),
            }
        )
    if not rows:
        return pd.DataFrame()
    frame = pd.DataFrame(rows).groupby(category_label, as_index=False)["排放量"].sum()
    frame = frame.sort_values("排放量", ascending=False)
    frame["完整值"] = frame["排放量"].map(lambda value: format_full_precision(value))
    return frame


def _render_record_detail(row: dict[str, Any] | None) -> None:
    st.markdown("#### 记录详情")
    if not row:
        st.info("请选择一条记录查看详情。切换记录不会重新核算。")
        return
    for label, value in record_detail_display(row):
        _text_card(label, value or "—")
    with st.expander("高级详情"):
        st.json(
            {
                key: row.get(key)
                for key in (
                    "Record_ID",
                    "Activity_Method",
                    "Governance_QC",
                    "Calculation_QC",
                    "Source_Sheet",
                    "Source_Row",
                    "Attribute_Match_Method",
                    "Attribute_Match_Status",
                    "Attribute_Source_File",
                    "Attribute_Source_Sheet",
                    "Attribute_Source_Row",
                    "Blocking_Codes",
                    "Warning_Codes",
                )
                if key in row
            }
        )


def _kpi_totals(rows: list[dict[str, Any]]) -> dict[str, str]:
    from decimal import Decimal

    activity = sum((Decimal(str(row["Activity_Data_kg"])) for row in rows if row.get("Activity_Data_kg")), Decimal("0"))
    emission_kg = sum((Decimal(str(row["Emission_kgCO2e"])) for row in rows if row.get("Emission_kgCO2e")), Decimal("0"))
    emission_t = sum((Decimal(str(row["Emission_tCO2e"])) for row in rows if row.get("Emission_tCO2e")), Decimal("0"))
    return {
        "activity_kg": format(activity, "f"),
        "emission_kgco2e": format(emission_kg, "f"),
        "emission_tco2e": format(emission_t, "f"),
    }


def _render_results_page(paths) -> None:
    st.subheader("核算结果与分析")
    e2e = _bound_e2e_view()
    if e2e is None:
        st.info("暂无数据")
        st.caption("本页只显示当前上传文件的核算结果，不会读取其他年度或历史运行。")
        return
    summary = e2e["e2e_summary"]
    canonical = e2e.get("canonical") or []
    st.caption(
        f"当前文件：{summary.get('Input_File') or st.session_state.get('Uploaded_File_Name') or '已恢复的当前运行'}"
    )
    unit_options = detected_unit_options(canonical, summary.get("Detected_Business_Units") or [])
    current_unit = st.session_state.get("selected_business_unit") or "全部"
    if current_unit not in unit_options:
        current_unit = "全部"
    selected_unit = st.selectbox("事业部", unit_options, index=unit_options.index(current_unit), key="results_unit")
    st.session_state["selected_business_unit"] = selected_unit
    canonical = filter_by_business_unit(canonical, selected_unit)
    totals = _kpi_totals(canonical)
    activity = totals.get("activity_kg")
    emission_t = totals.get("emission_tco2e")
    emission_kg = totals.get("emission_kgco2e")
    ef_value = next((row.get("EF_Value") for row in canonical if row.get("EF_Value")), summary.get("Factor_Value"))
    cols = st.columns(4)
    cols[0].metric(
        "总活动量",
        format_activity_display(activity),
        help=f"完整值：{format_full_precision(activity)} kg/year",
    )
    cols[1].metric(
        "排放因子",
        f"{ef_value} kgCO2e/kg" if ef_value else "—",
        help=f"完整值：{format_full_precision(ef_value)}",
    )
    cols[2].metric(
        "碳排放总量",
        format_emission_display(emission_t or emission_kg),
        help=f"完整值：{format_full_precision(emission_t or emission_kg)} tCO2e",
    )
    cols[3].metric("记录数", len(canonical))
    _text_card("核算状态", display_status(summary.get("Status")))
    _text_card("排放因子来源", display_route(summary.get("Factor_Route")))
    _text_card("当前文件", summary.get("Input_File") or st.session_state.get("Uploaded_File_Name"))
    files = st.session_state.get("Uploaded_Files") or []
    if len(files) > 1:
        _text_card("已上传文件", "；".join(str(item.get("name") or "") for item in files))
        note = summary.get("Attribute_Match_Note")
        if note:
            st.info(note)
    st.caption("因子用途：历史模拟　生产使用：否　结果全部来自本次上传")

    improvement_policy = comparison_ef_for_current_run(str(summary.get("Input_SHA256") or ""))
    if improvement_policy and canonical:
        improvement = factor_improvement_from_canonical(
            canonical,
            comparison_ef=improvement_policy["comparison_ef"],
        )
        if improvement:
            st.markdown("#### 因子变化模拟")
            st.info("保持当前文件活动量不变，仅替换排放因子。这不是年度同比，也不是供应商实际减排。")
            sim_cols = st.columns(4)
            sim_cols[0].metric(
                "2025按历史因子模拟排放",
                format_emission_display(improvement["simulated_emission_kgco2e"], unit="kgCO2e"),
                help=f"完整值：{format_full_precision(improvement['simulated_emission_kgco2e'])} kgCO2e",
            )
            sim_cols[1].metric(
                "2025当前因子排放",
                format_emission_display(improvement["current_emission_kgco2e"], unit="kgCO2e"),
                help=f"完整值：{format_full_precision(improvement['current_emission_kgco2e'])} kgCO2e",
            )
            sim_cols[2].metric(
                "因子变化带来的模拟减排量",
                format_emission_display(improvement["reduction_tco2e"]),
                help=f"完整值：{format_full_precision(improvement['reduction_tco2e'])} tCO2e",
            )
            sim_cols[3].metric(
                "排放因子下降比例",
                format_percentage_display(improvement["ef_decline_percent"], places=2),
                help=f"完整值：{format_full_precision(improvement['ef_decline_percent'])}%",
            )

    chemistry_frame = _chart_frame(canonical, category_field="Chemistry", category_label="化学体系")
    supplier_frame = _chart_frame(canonical, category_field="Supplier", category_label="供应商")
    if chemistry_frame.empty:
        chemistry_frame = _chart_frame(canonical, category_field="Business_Unit", category_label="事业部")
    if supplier_frame.empty:
        supplier_frame = _chart_frame(canonical, category_field="Purchase_Category", category_label="物料类别")
    left, right = st.columns(2)
    with left:
        st.markdown("#### 排放贡献")
        if chemistry_frame.empty:
            st.info("暂无数据")
        else:
            category = chemistry_frame.columns[0]
            order = (
                chemistry_display_order(chemistry_frame[category])
                if category == "化学体系"
                else chemistry_frame[category].tolist()
            )
            st.altair_chart(
                _summary_bar_chart(chemistry_frame, category=category, category_order=order),
                width="stretch",
            )
    with right:
        st.markdown("#### 分类汇总")
        if supplier_frame.empty:
            st.info("暂无数据")
        else:
            category = supplier_frame.columns[0]
            st.altair_chart(
                _summary_bar_chart(
                    supplier_frame,
                    category=category,
                    category_order=supplier_frame[category].tolist(),
                ),
                width="stretch",
            )

    st.markdown("#### 逐条核算结果")
    if not canonical:
        st.info("暂无数据")
    else:
        show = pd.DataFrame(canonical).rename(
            columns={
                "Record_ID": "记录编号",
                "Business_Unit": "事业部",
                "Product_Description": "物料描述",
                "Activity_Data_kg": "活动量",
                "EF_Value": "排放因子",
                "EF_Source": "排放因子来源",
                "Emission_tCO2e": "排放量",
                "Calculation_QC": "核算状态",
                "Warning_Codes": "异常提示",
            }
        )
        keep = [
            item
            for item in ("记录编号", "事业部", "物料描述", "活动量", "排放因子", "排放因子来源", "排放量", "核算状态", "异常提示")
            if item in show.columns
        ]
        if "核算状态" in show.columns:
            show["核算状态"] = show["核算状态"].map(display_status)
        if "异常提示" in show.columns:
            show["异常提示"] = show["异常提示"].map(display_reason_code)
        st.dataframe(show[keep], hide_index=True, width="stretch")
        record_ids = [str(row.get("Record_ID") or "") for row in canonical if row.get("Record_ID")]
        previous = st.session_state.get("selected_record_id")
        default_index = record_ids.index(previous) if previous in record_ids else 0
        selected_id = st.selectbox("选择记录编号", record_ids, index=default_index, key="record_picker")
        st.session_state["selected_record_id"] = selected_id
        selected_row = next((row for row in canonical if str(row.get("Record_ID")) == selected_id), None)
        _render_record_detail(selected_row)

    run_dir = _active_run_dir()
    snapshot = load_run_snapshot(run_dir, paths) if run_dir else {"technical_status": {}}
    technical = snapshot["technical_status"]
    if "day7" not in technical:
        with st.expander("高级：分步启动碳核算"):
            mode_label = st.radio(
                "排放因子来源",
                ["使用历史模拟因子", "上传因子文件"],
                horizontal=True,
                key="factor_mode",
            )
            factor_path = None
            if mode_label == "上传因子文件":
                uploaded = st.file_uploader("上传因子文件", type=["csv", "xlsx"], key="factor_upload")
                if uploaded:
                    factor_path = save_uploaded_file(uploaded.getvalue(), uploaded.name, paths)
            accepted = st.checkbox("我了解质量提示和历史模拟结果不会被自动关闭", key="factor_warning_confirmation")
            mode = "historical_simulation" if mode_label == "使用历史模拟因子" else "uploaded_factor"
            can_run = (
                technical.get("day5") == "PASS"
                and "day6" not in technical
                and accepted
                and (mode == "historical_simulation" or factor_path is not None)
            )
            if st.button("开始电芯碳核算并生成结果文件", disabled=not can_run):
                _execute(
                    "排放因子匹配、碳核算与结果文件生成",
                    lambda: run_calculation_and_export(
                        run_dir,
                        paths,
                        factor_mode=mode,
                        factor_input_path=factor_path,
                    ),
                )


def _render_download_page(paths) -> None:
    st.subheader("结果下载")
    e2e = _bound_e2e_view()
    if e2e is None:
        st.info("暂无数据")
        st.caption("请先完成当前文件核算后再下载。本页不会提供其他运行的结果文件。")
        with st.expander("高级 / 审计文件"):
            st.write("当前没有可下载的审计文件。")
        return
    run_dir = _active_run_dir()
    pack = build_business_download_pack(run_dir, st.session_state.get("selected_business_unit") or "全部")
    st.caption("当前文件的下载均来自本次运行，并按当前选择的事业部筛选。")
    buttons = st.columns(2)
    keys = ("cell_detail", "carbon_result", "third_party", "package")
    for index, key in enumerate(keys):
        item = pack[key]
        with buttons[index % 2]:
            st.download_button(
                item["display_name"],
                data=item["data"],
                file_name=item["download_name"],
                mime=item["mime"],
                key=f"biz-dl-{key}",
                width="stretch",
            )
    download_dir = e2e.get("download_dir")
    with st.expander("高级 / 审计文件"):
        if download_dir and (Path(download_dir) / "download_manifest.json").is_file():
            artifacts = wp6_8_download_artifacts(Path(download_dir))
            for item in artifacts:
                st.download_button(
                    item["download_name"],
                    data=item["data"],
                    file_name=item["download_name"],
                    mime=item.get("mime"),
                    key=f"audit-dl-{item['download_name']}",
                )
        else:
            st.write("暂无数据")
        historical = latest_wp6_8_run(paths.project_root.parent / "wp6-8" / "runs")
        if historical is not None:
            st.caption("以下为历史正式整合证据，不是本次上传结果。")
            view = load_wp6_8_view(historical)
            if view.get("ledger"):
                st.dataframe(pd.DataFrame(view["ledger"]).head(5), hide_index=True, width="stretch")


_init_state()
st.title("电芯数据碳核算程序")

try:
    day9_paths = load_day9_paths(DEFAULT_LOCAL_CONFIG)
except Exception as config_error:
    st.error("无法读取运行配置。")
    st.json(safe_error(config_error))
    st.stop()

_restore_completed_current_run(day9_paths)
if st.session_state.get("current_run_restore_warning"):
    st.warning(st.session_state["current_run_restore_warning"])

readiness = day9_paths.readiness()
if readiness["status"] == "PASS":
    st.sidebar.success("运行环境已就绪")
else:
    st.sidebar.error("运行环境未就绪")

with st.sidebar:
    page = st.radio(
        "功能导航",
        PAGES,
        key="business_page",
    )
    if st.button("运行说明", width="stretch"):
        _operation_dialog()
    with st.expander("高级工具"):
        st.caption("开发与审计信息，普通核算不必使用。")
        display_paths = path_config_for_display(day9_paths)
        st.write(
            {
                "运行目录": display_paths["run_root"],
                "结果目录": display_paths["output_root"],
            }
        )
st.divider()

if page == "数据导入与识别":
    _render_collection_page(day9_paths)
elif page == "数据能力与核算范围":
    _render_mapping_page(day9_paths)
elif page == "数据质量与异常":
    _render_cleaning_page(day9_paths)
elif page == "核算结果与分析":
    _render_results_page(day9_paths)
else:
    _render_download_page(day9_paths)
