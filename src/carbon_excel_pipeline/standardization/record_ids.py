"""Stable, profile-isolated record identifier generation."""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from typing import Any


FORMAL_KEY_FIELDS = ("Source_File", "Source_Sheet", "Source_Row")


def source_key(record: dict[str, Any]) -> tuple[str, str, int]:
    return (
        str(record["Source_File"]),
        str(record["Source_Sheet"]),
        int(record["Source_Row"]),
    )


def stable_identity(record: dict[str, Any], profile_id: str) -> str:
    payload = {
        "profile_id": profile_id,
        "source_file": str(record.get("Source_File", "")),
        "source_sheet": str(record.get("Source_Sheet", "")),
        "source_row": int(record.get("Source_Row", 0)),
        "description": str(record.get("Product_Description_Raw", "")),
        "pcs": str(record.get("PCS_Clean", "")),
        "unit_weight": str(record.get("Unit_Weight_Clean", "")),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


@dataclass(slots=True)
class RecordIdResolver:
    profile: dict[str, Any]
    mapping_config: dict[str, Any]
    frozen_ids: dict[tuple[str, str, int], str] = field(default_factory=dict)
    seen_generated: dict[str, str] = field(default_factory=dict)

    def resolve(self, record: dict[str, Any]) -> tuple[str, str]:
        profile_id = self.profile["profile_id"]
        identity = stable_identity(record, profile_id)
        if profile_id != "public_synthetic_profile":
            raise ValueError(f"Unsupported profile_id: {profile_id}")
        settings = self.mapping_config["record_id"]
        digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()
        modulus = 10 ** int(settings["public_digits"])
        number = int(digest[:16], 16) % modulus
        record_id = f"{settings['public_namespace']}{number:0{settings['public_digits']}d}"
        if not re.fullmatch(self.profile["record_id_regex"], record_id):
            raise ValueError("Generated public Record_ID violates the profile regex.")
        return self._register(record_id, identity), "PUBLIC_STABLE_HASH"

    def _register(self, record_id: str, identity: str) -> str:
        previous = self.seen_generated.get(record_id)
        if previous is not None and previous != identity:
            raise ValueError(
                "Stable Record_ID hash collision; change the configured namespace rule."
            )
        self.seen_generated[record_id] = identity
        return record_id
