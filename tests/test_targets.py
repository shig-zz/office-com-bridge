from __future__ import annotations

from office_bridge.targets import (
    MAX_RECORDS,
    TargetRegistry,
    TargetState,
    dedupe_key,
    new_target_id,
)


def add_book(
    registry: TargetRegistry,
    name: str = "Book1.xlsx",
    full: str = r"C:\docs\Book1.xlsx",
    pid: int = 100,
    hwnd: int = 1000,
):
    return registry.add(
        app_type="excel",
        display_name=name,
        dedupe_key=dedupe_key("excel", pid, full, hwnd),
        metadata={"full_name": full, "path": r"C:\docs"},
        app_metadata={"hwnd": hwnd, "pid": pid, "application": "Excel"},
    )


def test_new_target_id_format_and_uniqueness() -> None:
    ids = {new_target_id() for _ in range(500)}
    assert len(ids) == 500
    for target_id in ids:
        assert target_id.startswith("t-")
        assert len(target_id) == 14


def test_dedupe_key_saved_document_uses_pid_and_lowercase_path() -> None:
    a = dedupe_key("excel", 100, r"C:\Docs\BOOK1.xlsx", 1000)
    b = dedupe_key("excel", 100, r"c:\docs\book1.xlsx", 2000)
    assert a == b
    assert dedupe_key("excel", 101, r"C:\docs\book1.xlsx", 1000) != a


def test_dedupe_key_unsaved_document_falls_back_to_hwnd() -> None:
    assert dedupe_key("word", 100, "", 555) == dedupe_key("word", 100, "", 555)
    assert dedupe_key("word", 100, "", 555) != dedupe_key("word", 100, "", 556)


def test_capture_dedupe_reuses_existing_record() -> None:
    registry = TargetRegistry()
    first = add_book(registry)

    # BridgeCore's capture flow: find by dedupe key, only add when missing.
    key = first.dedupe_key
    existing = registry.find_active_by_dedupe(key)
    assert existing is first
    assert registry.find_active_by_dedupe(key) is first  # stable across calls

    other = add_book(registry, name="Book2.xlsx", full=r"C:\docs\Book2.xlsx", hwnd=2000)
    assert other.target_id != first.target_id
    assert len(registry) == 2


def test_select_and_selected() -> None:
    registry = TargetRegistry()
    first = add_book(registry)
    other = add_book(registry, name="Book2.xlsx", full=r"C:\docs\Book2.xlsx", hwnd=2000)

    assert registry.selected_target_id is None
    registry.select(other.target_id)
    assert registry.selected_target_id == other.target_id
    assert registry.selected() is other
    assert other.last_selected_at is not None

    assert registry.select("t-doesnotexist") is None
    registry.clear_selection()
    assert registry.selected_target_id is None


def test_expire_marks_state_and_clears_selection() -> None:
    registry = TargetRegistry()
    first = add_book(registry)
    registry.select(first.target_id)

    record = registry.expire(first.target_id, "com_reference_lost")
    assert record is not None
    assert record.state is TargetState.EXPIRED
    assert record.expired_reason == "com_reference_lost"
    assert registry.selected_target_id is None
    assert registry.find_active_by_dedupe(first.dedupe_key) is None
    # re-capture after expiry creates a fresh target
    fresh = add_book(registry)
    assert fresh.target_id != first.target_id


def test_remove_clears_selection_and_returns_record() -> None:
    registry = TargetRegistry()
    first = add_book(registry)
    registry.select(first.target_id)
    removed = registry.remove(first.target_id)
    assert removed is not None
    assert registry.get(first.target_id) is None
    assert registry.selected_target_id is None
    assert registry.remove(first.target_id) is None


def test_update_metadata_refreshes_fields() -> None:
    registry = TargetRegistry()
    first = add_book(registry)
    record = registry.update_metadata(
        first.target_id,
        display_name="Renamed.xlsx",
        metadata={"full_name": r"C:\docs\Renamed.xlsx", "path": r"C:\docs"},
        app_metadata={"hwnd": 1001, "pid": 100, "application": "Excel"},
    )
    assert record is not None
    assert record.display_name == "Renamed.xlsx"
    assert record.app_metadata["hwnd"] == 1001
    assert record.last_refreshed_at is not None


def test_purge_expired_removes_oldest_expired_only() -> None:
    registry = TargetRegistry()
    stale = [add_book(registry, full=f"C:\\docs\\S{i}.xlsx", hwnd=i) for i in range(3)]
    keep = add_book(registry, full=r"C:\docs\Keep.xlsx", hwnd=9999)
    for record in stale:
        registry.expire(record.target_id, "closed")

    assert registry.purge_expired(max_records=5) == []
    removed = registry.purge_expired(max_records=3)
    assert removed == [stale[0].target_id]
    assert len(registry) == 3
    removed = registry.purge_expired(max_records=1)
    assert removed == [r.target_id for r in stale[1:]]
    assert registry.get(keep.target_id) is keep


def test_purge_never_evicts_active_records() -> None:
    registry = TargetRegistry()
    for index in range(MAX_RECORDS + 5):
        add_book(registry, full=f"C:\\docs\\B{index}.xlsx", hwnd=index)
    removed = registry.purge_expired()
    assert removed == []
    assert len(registry) == MAX_RECORDS + 5


def test_public_and_summary_shape() -> None:
    registry = TargetRegistry()
    record = add_book(registry)
    registry.select(record.target_id)

    public = record.public()
    assert public["target_id"] == record.target_id
    assert public["app_type"] == "excel"
    assert public["display_name"] == "Book1.xlsx"
    assert public["state"] == "active"
    assert public["metadata"]["full_name"] == r"C:\docs\Book1.xlsx"
    assert public["app_metadata"]["pid"] == 100

    summary = record.summary()
    assert summary["target_id"] == record.target_id
    assert summary["application"] == "Excel"
    assert summary["document"] == "Book1.xlsx"
    assert summary["hwnd"] == 1000

    listing = registry.list_public()
    assert len(listing) == 1
    assert listing[0]["target_id"] == record.target_id
