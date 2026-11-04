"""Target registry: multiple Office documents identified by stable target IDs.

This module is intentionally free of COM and window-API imports. It holds
app-agnostic ``TargetRecord`` metadata; Office-specific data (hwnd, pid, sheet
names, ...) lives in ``app_metadata``. The COM object references themselves are
owned elsewhere (``BridgeCore._com_refs``) and never exposed here.

All registry access must stay on the COM worker thread; ``BridgeCore`` routes
every access through ``ComWorker.call``, so no locking is needed here.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from typing import Any

# Expired records are kept for audit until the registry exceeds this size;
# active records are never evicted automatically.
MAX_RECORDS = 50

_TARGET_ID_PREFIX = "t-"


class TargetState(str, Enum):
    ACTIVE = "active"
    EXPIRED = "expired"


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def new_target_id() -> str:
    """Opaque ID; never derived from hwnd, pid, or file names."""
    return _TARGET_ID_PREFIX + uuid.uuid4().hex[:12]


def dedupe_key(app_type: str, pid: int, full_name: str, hwnd: int) -> str:
    """Key used to detect duplicate capture of the same open document.

    Saved documents key on (pid, lowercased full path); unsaved documents
    (empty full name) fall back to the window handle.
    """
    if full_name:
        return f"{int(pid)}:{full_name.strip().lower()}"
    return f"hwnd:{int(hwnd)}"


@dataclass(slots=True)
class TargetRecord:
    target_id: str
    app_type: str  # "word" | "excel" | "powerpoint" (lowercase)
    display_name: str  # document file name
    dedupe_key: str
    state: TargetState = TargetState.ACTIVE
    expired_reason: str | None = None  # closed | window_gone | com_reference_lost
    created_at: str = field(default_factory=utc_now)
    last_selected_at: str | None = None
    last_refreshed_at: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)  # common: full_name, path
    app_metadata: dict[str, Any] = field(default_factory=dict)  # office: hwnd, pid, ...

    def public(self) -> dict[str, Any]:
        return {
            "target_id": self.target_id,
            "app_type": self.app_type,
            "display_name": self.display_name,
            "state": self.state.value,
            "expired_reason": self.expired_reason,
            "created_at": self.created_at,
            "last_selected_at": self.last_selected_at,
            "last_refreshed_at": self.last_refreshed_at,
            "metadata": dict(self.metadata),
            "app_metadata": dict(self.app_metadata),
        }

    def summary(self) -> dict[str, Any]:
        return {
            "target_id": self.target_id,
            "application": self.app_metadata.get("application", self.app_type),
            "document": self.display_name,
            "full_name": self.metadata.get("full_name", ""),
            "hwnd": self.app_metadata.get("hwnd", 0),
            "pid": self.app_metadata.get("pid", 0),
        }


class TargetRegistry:
    def __init__(self) -> None:
        self._records: dict[str, TargetRecord] = {}
        self.selected_target_id: str | None = None

    def __len__(self) -> int:
        return len(self._records)

    def get(self, target_id: str) -> TargetRecord | None:
        return self._records.get(target_id)

    def records(self) -> list[TargetRecord]:
        return list(self._records.values())

    def active_records(self) -> list[TargetRecord]:
        return [r for r in self._records.values() if r.state is TargetState.ACTIVE]

    def find_active_by_dedupe(self, key: str) -> TargetRecord | None:
        for record in self._records.values():
            if record.state is TargetState.ACTIVE and record.dedupe_key == key:
                return record
        return None

    def add(
        self,
        *,
        app_type: str,
        display_name: str,
        dedupe_key: str,
        metadata: dict[str, Any],
        app_metadata: dict[str, Any],
        now: str | None = None,
    ) -> TargetRecord:
        now = now or utc_now()
        target_id = new_target_id()
        while target_id in self._records:
            target_id = new_target_id()
        record = TargetRecord(
            target_id=target_id,
            app_type=app_type,
            display_name=display_name,
            dedupe_key=dedupe_key,
            created_at=now,
            metadata=dict(metadata),
            app_metadata=dict(app_metadata),
        )
        self._records[target_id] = record
        return record

    def update_metadata(
        self,
        target_id: str,
        *,
        display_name: str,
        metadata: dict[str, Any],
        app_metadata: dict[str, Any],
        now: str | None = None,
    ) -> TargetRecord | None:
        record = self._records.get(target_id)
        if record is None:
            return None
        record.display_name = display_name
        record.metadata = dict(metadata)
        record.app_metadata = dict(app_metadata)
        record.last_refreshed_at = now or utc_now()
        return record

    def select(self, target_id: str, now: str | None = None) -> TargetRecord | None:
        record = self._records.get(target_id)
        if record is None:
            return None
        self.selected_target_id = target_id
        record.last_selected_at = now or utc_now()
        return record

    def clear_selection(self) -> None:
        self.selected_target_id = None

    def selected(self) -> TargetRecord | None:
        if self.selected_target_id is None:
            return None
        return self._records.get(self.selected_target_id)

    def expire(self, target_id: str, reason: str, now: str | None = None) -> TargetRecord | None:
        record = self._records.get(target_id)
        if record is None:
            return None
        record.state = TargetState.EXPIRED
        record.expired_reason = reason
        record.last_refreshed_at = now or utc_now()
        if self.selected_target_id == target_id:
            self.selected_target_id = None
        return record

    def remove(self, target_id: str) -> TargetRecord | None:
        record = self._records.pop(target_id, None)
        if record is None:
            return None
        if self.selected_target_id == target_id:
            self.selected_target_id = None
        return record

    def purge_expired(self, max_records: int = MAX_RECORDS) -> list[str]:
        """Drop the oldest expired records until the registry fits max_records.

        Active records are never removed; if active records alone exceed the
        limit, the registry is allowed to stay over it.
        """
        removed: list[str] = []
        if len(self._records) <= max_records:
            return removed
        for target_id, record in list(self._records.items()):
            if len(self._records) <= max_records:
                break
            if record.state is TargetState.EXPIRED:
                del self._records[target_id]
                removed.append(target_id)
        return removed

    def list_public(self) -> list[dict[str, Any]]:
        return [record.public() for record in self._records.values()]

    def clear(self) -> None:
        self._records.clear()
        self.selected_target_id = None
