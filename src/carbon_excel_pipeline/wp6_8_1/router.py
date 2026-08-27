"""Independent Activity, Factor and Boundary routing. Year is never the selector."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Mapping


PROJECT_ROOT = Path(__file__).resolve().parents[3]
DEFAULT_CATALOG = PROJECT_ROOT / "config" / "wp6" / "wp6_8_1_route_policies.json"

ACTIVITY_DIRECT = "DIRECT_REPORTED_MASS"
ACTIVITY_PCS = "PCS_WEIGHT_DERIVED"
FACTOR_EMBEDDED = "SOURCE_EMBEDDED_FACTOR"
FACTOR_SIMULATION = "HISTORICAL_SIMULATION_FACTOR"
FACTOR_EXTERNAL = "EXTERNAL_FACTOR_RESULT"
FACTOR_MISSING = "FACTOR_NOT_AVAILABLE"
BOUNDARY_MISSING = "BOUNDARY_POLICY_NOT_AVAILABLE"


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _truth(value: Any) -> bool:
    return _text(value).upper() in {"TRUE", "1", "YES", "PASS"}


def _sha(value: Any) -> str:
    return _text(value).upper()


def _supported_activity_routes(capability: Mapping[str, Any]) -> list[str]:
    routes: list[str] = []
    if int(capability.get("direct_reported_mass_count") or 0) > 0:
        routes.append(ACTIVITY_DIRECT)
    if int(capability.get("pcs_weight_derived_count") or 0) > 0:
        routes.append(ACTIVITY_PCS)
    explicit = capability.get("supported_activity_paths")
    if isinstance(explicit, list):
        for item in explicit:
            route = _text(item)
            if route and route not in routes:
                routes.append(route)
    return routes


def _embedded_factor_ready(capability: Mapping[str, Any]) -> bool:
    if "factor_ready_count" in capability:
        return int(capability.get("factor_ready_count") or 0) > 0
    return _truth(capability.get("factor_ready"))


def _sha_allowed(policy: Mapping[str, Any], input_sha256: str) -> bool:
    required = [_sha(item) for item in policy.get("match", {}).get("input_sha256", [])]
    if not required:
        return True
    return _sha(input_sha256) in required


def _activity_allowed(policy: Mapping[str, Any], activity_route: str | None) -> bool:
    required = _text(policy.get("match", {}).get("required_activity_route"))
    if not required:
        return True
    return activity_route == required


def decide_processing_routes(
    *,
    capability: Mapping[str, Any],
    input_sha256: str,
    catalog: Mapping[str, Any] | None = None,
    catalog_path: Path | None = None,
    requested_activity_route: str | None = None,
) -> dict[str, Any]:
    """Return three independent routing facts plus an overall routing status."""

    payload = catalog or _load_json(catalog_path or DEFAULT_CATALOG)
    supported = _supported_activity_routes(capability)
    warnings: list[str] = []
    blocking: list[str] = []

    activity_route: str | None = None
    activity_reason = ""
    if requested_activity_route:
        if requested_activity_route not in supported:
            blocking.append("REQUESTED_ACTIVITY_PATH_NOT_SUPPORTED")
            activity_reason = "用户请求的活动路径当前数据并不支持。"
        else:
            activity_route = requested_activity_route
            activity_reason = "显式 Processing Policy 在 Capability 支持集合中选择了活动路径。"
    elif not supported:
        blocking.append("NO_ACTIVITY_PATH")
        activity_reason = "当前识别结果不支持 PCS×单重 或 直接年度质量 任一活动路径。"
    elif len(supported) == 1:
        activity_route = supported[0]
        activity_reason = "Capability 仅支持一条活动路径，Router 未使用年份。"
    else:
        blocking.append("MULTIPLE_VALID_ROUTES")
        activity_reason = "同时支持多条活动路径，且没有显式 Policy，禁止静默优先 PCS。"

    embedded = _embedded_factor_ready(capability)
    factor_policy: dict[str, Any] | None = None
    for policy in payload.get("factor_policies", []):
        match = policy.get("match", {})
        needs_embedded = bool(match.get("require_embedded_factor"))
        if needs_embedded != embedded:
            continue
        if not _sha_allowed(policy, input_sha256):
            continue
        if not _activity_allowed(policy, activity_route):
            continue
        factor_policy = dict(policy)
        break

    if factor_policy is None:
        factor_route = FACTOR_MISSING
        factor_reason = "没有匹配的受控因子 Policy；禁止把 1.250000 当作所有 PCS 文件的默认因子。"
        factor_ready = False
    else:
        factor_route = _text(factor_policy.get("factor_route")) or FACTOR_MISSING
        factor_reason = f"匹配受控因子 Policy：{factor_policy.get('policy_id')}。"
        factor_ready = factor_route != FACTOR_MISSING

    boundary_policy: dict[str, Any] | None = None
    for policy in payload.get("boundary_policies", []):
        match = policy.get("match", {})
        if match.get("require_embedded_factor") and not embedded:
            continue
        if match.get("require_embedded_factor") is False and embedded:
            continue
        if not _sha_allowed(policy, input_sha256):
            continue
        if not _activity_allowed(policy, activity_route):
            continue
        boundary_policy = dict(policy)
        break

    if boundary_policy is None:
        boundary_id = BOUNDARY_MISSING
        boundary_reason = "没有匹配的受控边界 Policy；禁止把已有合成示例范围套用到未知文件。"
        boundary_ready = False
        processor = None
    else:
        boundary_id = _text(boundary_policy.get("policy_id"))
        boundary_reason = f"匹配受控边界 Policy：{boundary_id}。"
        boundary_ready = True
        processor = boundary_policy.get("processor")

    activity_ready = activity_route is not None
    emission_ready = activity_ready and factor_ready and boundary_ready
    if blocking and "MULTIPLE_VALID_ROUTES" in blocking:
        status = "MULTIPLE_VALID_ROUTES"
    elif not activity_ready:
        status = "BLOCKED"
    elif not emission_ready:
        status = "PARTIAL_RESULT"
        if not boundary_ready:
            warnings.append("BOUNDARY_POLICY_NOT_AVAILABLE")
        if not factor_ready:
            warnings.append("FACTOR_POLICY_NOT_AVAILABLE")
    else:
        status = "ROUTED"

    return {
        "schema_version": "WP6_8_1_ROUTE_DECISION_V1",
        "status": status,
        "Input_SHA256": _sha(input_sha256),
        "Supported_Activity_Paths": supported,
        "Activity_Route": activity_route,
        "Activity_Route_Reason": activity_reason,
        "Activity_Ready": activity_ready,
        "Factor_Route": factor_route,
        "Factor_Policy_ID": None if factor_policy is None else factor_policy.get("policy_id"),
        "Factor_Value": None if factor_policy is None else factor_policy.get("factor_value"),
        "Factor_Unit": None if factor_policy is None else factor_policy.get("factor_unit"),
        "Factor_Source": None if factor_policy is None else factor_policy.get("factor_source"),
        "Factor_Usage": None if factor_policy is None else factor_policy.get("factor_usage"),
        "Simulation_Flag": True if factor_policy is None else bool(factor_policy.get("simulation_flag", True)),
        "Production_Eligible": False if factor_policy is None else bool(factor_policy.get("production_eligible", False)),
        "Factor_Route_Reason": factor_reason,
        "Factor_Ready": factor_ready,
        "Boundary_Policy": boundary_id,
        "Boundary_Policy_Reason": boundary_reason,
        "Boundary_Ready": boundary_ready,
        "Boundary_Config_Path": None if boundary_policy is None else boundary_policy.get("config_path"),
        "Factor_Config_Path": None if factor_policy is None else factor_policy.get("config_path"),
        "Processor": processor,
        "Emission_Ready": emission_ready,
        "Warnings": warnings,
        "Blocking_Reasons": blocking,
        "Year_Used_As_Router": False,
    }
