"""Explicit processing policy kept separate from capability detection."""

from __future__ import annotations

from collections.abc import Iterable

from carbon_excel_pipeline.errors import PipelineUserError

from .models import ActivityPath


def select_activity_path(
    supported_paths: Iterable[ActivityPath | str],
    *,
    requested_path: ActivityPath | str,
) -> ActivityPath:
    """Select a requested path only when the detector reported it as supported."""

    supported = {ActivityPath(item) for item in supported_paths}
    requested = ActivityPath(requested_path)
    if requested not in supported:
        raise PipelineUserError(
            stage="PROCESSING_POLICY",
            error_code="REQUESTED_ACTIVITY_PATH_NOT_SUPPORTED",
            message_cn="当前处理策略选择了该记录不具备的活动数据路径。",
            source_location="Supported_Activity_Paths",
            original_value={
                "requested": requested,
                "supported": sorted(str(item) for item in supported),
            },
            rule="Capability只说明可用路径；Processing Policy必须显式选择其中一条。",
            impact="该记录不能进入正式活动数据转换。",
            fix_suggestion="修正处理策略或补齐该路径所需的源数据与单位。",
        )
    return requested
