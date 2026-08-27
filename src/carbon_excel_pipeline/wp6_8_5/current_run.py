"""Persist and strictly restore one completed Streamlit Current Run."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from carbon_excel_pipeline.io.excel_importer import sha256_file
from carbon_excel_pipeline.wp6_8_4.input_set import input_set_sha256
from carbon_excel_pipeline.wp6_8_6.record_ids import RECORD_ID_SCHEMA_VERSION


POINTER_NAME = "current_run.json"


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _pointer_path(run_root: Path) -> Path:
    return run_root.expanduser().resolve() / POINTER_NAME


def persist_current_run(run_dir: Path) -> dict[str, Any] | None:
    run = run_dir.expanduser().resolve()
    summary_path = run / "e2e_run_summary.json"
    stage_path = run / "pipeline_stage_status.json"
    canonical_path = run / "08_download" / "canonical_results.csv"
    if not (summary_path.is_file() and stage_path.is_file() and canonical_path.is_file()):
        return None
    summary = _load_json(summary_path)
    if summary.get("Record_ID_Schema_Version") != RECORD_ID_SCHEMA_VERSION:
        return None
    stages = _load_json(stage_path).get("stages") or {}
    if stages.get("COMPLETED") not in {"PASS", "PASS_WITH_WARNING", "PARTIAL_RESULT"}:
        return None
    run_root = run.parent
    files: list[dict[str, Any]] = []
    for item in summary.get("Input_Files") or []:
        name = str(item.get("name") or item.get("file_name") or "").strip()
        copy_path = run / "00_input_copy" / name
        files.append(
            {
                "name": name,
                "sha256": str(item.get("sha256") or item.get("Input_SHA256") or "").upper(),
                "role": item.get("role"),
                "path": str(copy_path.resolve()) if copy_path.is_file() else str(item.get("path") or ""),
            }
        )
    pointer = {
        "schema_version": "WP6_8_5_CURRENT_RUN_POINTER_V1",
        "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
        "Current_Run_ID": summary.get("Run_ID") or run.name,
        "Current_Run_Root": str(run),
        "Input_Set_SHA256": summary.get("Input_Set_SHA256"),
        "Input_Files": files,
        "Status": summary.get("Status"),
    }
    path = _pointer_path(run_root)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(pointer, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)
    return pointer


def restore_current_run(run_root: Path) -> dict[str, Any] | None:
    root = run_root.expanduser().resolve()
    path = _pointer_path(root)
    if not path.is_file():
        return None
    try:
        pointer = _load_json(path)
        run = Path(str(pointer.get("Current_Run_Root") or "")).resolve()
        if not _inside(run, root) or run.parent != root:
            return None
        summary_path = run / "e2e_run_summary.json"
        stage_path = run / "pipeline_stage_status.json"
        canonical_path = run / "08_download" / "canonical_results.csv"
        if not (summary_path.is_file() and stage_path.is_file() and canonical_path.is_file()):
            return None
        summary = _load_json(summary_path)
        if (
            pointer.get("Record_ID_Schema_Version") != RECORD_ID_SCHEMA_VERSION
            or summary.get("Record_ID_Schema_Version") != RECORD_ID_SCHEMA_VERSION
        ):
            return {
                "Restore_Status": "LEGACY_RECORD_ID_SCHEMA",
                "Record_ID_Schema_Version": str(
                    summary.get("Record_ID_Schema_Version") or "LEGACY_OR_MISSING"
                ),
                "Current_Run_ID": summary.get("Run_ID") or run.name,
                "Current_Run_Root": str(run),
                "Message": "检测到旧版记录编号运行结果，请使用当前输入重新运行。",
            }
        stages = _load_json(stage_path).get("stages") or {}
        if summary.get("Run_ID") != pointer.get("Current_Run_ID"):
            return None
        if stages.get("COMPLETED") not in {"PASS", "PASS_WITH_WARNING", "PARTIAL_RESULT"}:
            return None
        summary_files = summary.get("Input_Files") or []
        if input_set_sha256(summary_files) != str(pointer.get("Input_Set_SHA256") or "").upper():
            return None
        restored_files: list[dict[str, Any]] = []
        for item in pointer.get("Input_Files") or []:
            name = str(item.get("name") or "")
            copy_path = run / "00_input_copy" / name
            expected = str(item.get("sha256") or "").upper()
            if not copy_path.is_file() or sha256_file(copy_path) != expected:
                return None
            restored_files.append({**item, "path": str(copy_path.resolve())})
        primary = next((item for item in restored_files if item.get("role") == "主核算数据"), None)
        if primary is None:
            return None
        return {
            "Restore_Status": "PASS",
            "Record_ID_Schema_Version": RECORD_ID_SCHEMA_VERSION,
            "Current_Run_ID": pointer["Current_Run_ID"],
            "Current_Run_Root": str(run),
            "Input_Set_SHA256": str(pointer.get("Input_Set_SHA256") or "").upper(),
            "Uploaded_Files": restored_files,
            "Uploaded_File_Name": primary.get("name"),
            "Uploaded_File_SHA256": primary.get("sha256"),
            "Uploaded_File_Temp_Path": primary.get("path"),
            "Current_Source_Path": primary.get("path"),
        }
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return None


def clear_current_run_pointer(run_root: Path) -> None:
    path = _pointer_path(run_root)
    if path.is_file():
        path.unlink()
