"""Structured WP6-1 workbook recognition result models."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import StrEnum
from typing import Any


class RecognitionStatus(StrEnum):
    RECOGNIZED = "RECOGNIZED"
    RECOGNIZED_WITH_WARNING = "RECOGNIZED_WITH_WARNING"
    AMBIGUOUS = "AMBIGUOUS"
    UNRECOGNIZED = "UNRECOGNIZED"


class MatchType(StrEnum):
    EXACT = "EXACT"
    ALIAS = "ALIAS"
    UNMAPPED = "UNMAPPED"
    AMBIGUOUS = "AMBIGUOUS"


class MappingStatus(StrEnum):
    MAPPED = "MAPPED"
    UNMAPPED = "UNMAPPED"
    AMBIGUOUS = "AMBIGUOUS"


@dataclass(frozen=True)
class RecognitionWarning:
    code: str
    message_cn: str
    context: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class FieldMappingResult:
    column_index: int
    column_letter: str
    raw_header: str
    normalized_header: str
    semantic_field: str | None
    legacy_target_field: str | None
    match_type: MatchType
    detected_unit: str | None
    mapping_status: MappingStatus
    warning_code: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class HeaderCandidateResult:
    row_number: int
    score: int
    semantic_match_count: int
    unique_semantic_count: int
    nonempty_cell_count: int
    data_evidence_count: int
    mappings: list[FieldMappingResult]
    duplicate_semantics: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            **asdict(self),
            "mappings": [item.to_dict() for item in self.mappings],
        }


@dataclass
class SheetRecognitionResult:
    sheet_name: str
    sheet_index: int
    sheet_state: str
    physical_row_count: int
    column_count: int
    dimension: str
    formula_count: int
    merged_ranges: list[str]
    status: RecognitionStatus
    selected_header_row: int | None
    recognition_score: int
    recognized_field_count: int
    mappings: list[FieldMappingResult]
    header_candidates: list[HeaderCandidateResult]
    warnings: list[RecognitionWarning] = field(default_factory=list)

    @property
    def merged_cell_count(self) -> int:
        return len(self.merged_ranges)

    @property
    def recognized_fields(self) -> list[str]:
        return sorted(
            {
                item.semantic_field
                for item in self.mappings
                if item.semantic_field is not None
            }
        )

    @property
    def unmapped_fields(self) -> list[str]:
        return [
            item.raw_header
            for item in self.mappings
            if item.mapping_status == MappingStatus.UNMAPPED
        ]

    @property
    def duplicate_mappings(self) -> list[str]:
        return sorted(
            {
                item.semantic_field
                for item in self.mappings
                if item.mapping_status == MappingStatus.AMBIGUOUS
                and item.semantic_field is not None
            }
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet_name": self.sheet_name,
            "sheet_index": self.sheet_index,
            "sheet_state": self.sheet_state,
            "physical_row_count": self.physical_row_count,
            "column_count": self.column_count,
            "dimension": self.dimension,
            "formula_count": self.formula_count,
            "merged_cell_count": self.merged_cell_count,
            "merged_ranges": self.merged_ranges,
            "recognition_status": self.status,
            "selected_header_row": self.selected_header_row,
            "recognition_score": self.recognition_score,
            "recognized_field_count": self.recognized_field_count,
            "recognized_fields": self.recognized_fields,
            "unmapped_fields": self.unmapped_fields,
            "duplicate_mappings": self.duplicate_mappings,
            "mappings": [item.to_dict() for item in self.mappings],
            "header_candidates": [item.to_dict() for item in self.header_candidates],
            "warnings": [item.to_dict() for item in self.warnings],
        }


@dataclass
class WorkbookRecognitionResult:
    workbook_name: str
    input_fingerprint: str
    status: RecognitionStatus
    sheets: list[SheetRecognitionResult]
    best_candidate_sheet: str | None
    best_candidate_header_row: int | None
    warnings: list[RecognitionWarning] = field(default_factory=list)

    @property
    def sheet_count(self) -> int:
        return len(self.sheets)

    @property
    def best_candidate(self) -> SheetRecognitionResult | None:
        return next(
            (item for item in self.sheets if item.sheet_name == self.best_candidate_sheet),
            None,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "workbook_name": self.workbook_name,
            "input_fingerprint": self.input_fingerprint,
            "sheet_count": self.sheet_count,
            "recognition_status": self.status,
            "best_candidate_sheet": self.best_candidate_sheet,
            "best_candidate_header_row": self.best_candidate_header_row,
            "warnings": [item.to_dict() for item in self.warnings],
            "sheets": [item.to_dict() for item in self.sheets],
        }
