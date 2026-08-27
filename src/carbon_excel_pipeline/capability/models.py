"""Structured result models for WP6-2 capability decisions."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import StrEnum
from typing import Any


class ActivityPath(StrEnum):
    PCS_WEIGHT_DERIVED = "PCS_WEIGHT_DERIVED"
    DIRECT_REPORTED_MASS = "DIRECT_REPORTED_MASS"


class ValueStatus(StrEnum):
    VALID = "VALID"
    MISSING = "MISSING"
    NON_NUMERIC = "NON_NUMERIC"
    ZERO = "ZERO"
    NEGATIVE = "NEGATIVE"


class CapabilityStatus(StrEnum):
    CAPABLE = "CAPABLE"
    CAPABLE_WITH_WARNING = "CAPABLE_WITH_WARNING"
    PARTIALLY_CAPABLE = "PARTIALLY_CAPABLE"
    INCAPABLE = "INCAPABLE"


@dataclass(frozen=True, slots=True)
class PathDecision:
    path: ActivityPath
    supported: bool
    reason_codes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class RecordCapabilityResult:
    source_row: int
    status: CapabilityStatus
    supported_activity_paths: list[ActivityPath]
    activity_ready: bool
    factor_ready: bool
    factor_source_available: bool
    emission_ready: bool
    historical_result_available: bool
    historical_validation_ready: bool
    analysis_capabilities: dict[str, bool]
    value_statuses: dict[str, ValueStatus]
    path_decisions: list[PathDecision]
    warning_codes: list[str] = field(default_factory=list)
    blocking_codes: list[str] = field(default_factory=list)
    formula_fields: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "Source_Row": self.source_row,
            "Status": self.status,
            "Supported_Activity_Paths": self.supported_activity_paths,
            "Activity_Ready": self.activity_ready,
            "Factor_Ready": self.factor_ready,
            "Factor_Source_Available": self.factor_source_available,
            "Emission_Ready": self.emission_ready,
            "Historical_Result_Available": self.historical_result_available,
            "Historical_Validation_Ready": self.historical_validation_ready,
            "Analysis_Capabilities": self.analysis_capabilities,
            "Value_Statuses": self.value_statuses,
            "Path_Decisions": [item.to_dict() for item in self.path_decisions],
            "Warning_Codes": self.warning_codes,
            "Blocking_Codes": self.blocking_codes,
            "Formula_Fields": self.formula_fields,
        }
