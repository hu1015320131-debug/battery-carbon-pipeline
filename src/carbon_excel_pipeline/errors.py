"""Safe, structured errors for users and testable pipeline behavior."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(slots=True)
class PipelineUserError(Exception):
    stage: str
    error_code: str
    message_cn: str
    source_location: str
    original_value: Any
    rule: str
    impact: str
    fix_suggestion: str

    def __post_init__(self) -> None:
        Exception.__init__(self, self.message_cn)

    def to_dict(self) -> dict[str, Any]:
        return {
            "stage": self.stage,
            "error_code": self.error_code,
            "message_cn": self.message_cn,
            "source_location": self.source_location,
            "original_value": self.original_value,
            "rule": self.rule,
            "impact": self.impact,
            "fix_suggestion": self.fix_suggestion,
        }
