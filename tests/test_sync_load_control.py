import sqlite3
from unittest.mock import Mock

import sync_utils


class FakeResponse:
    def __init__(self, status_code=200, payload=None):
        self.status_code = status_code
        self._payload = payload if payload is not None else []

    def json(self):
        return self._payload


class FakeWCAPI:
    def __init__(self):
        self.paths = []

    def get(self, path, params=None):
        self.paths.append((path, params))
        order_id = int(path.split("/")[1])
        return FakeResponse(
            payload=[
                {
                    "id": order_id * 10,
                    "note": f"note-{order_id}",
                    "date_created": "2026-07-31T10:00:00",
                    "customer_note": False,
                    "author": "WooCommerce",
                    "added_by_user": False,
                }
            ]
        )


def make_note_db():
    conn = sqlite3.connect(":memory:")
    conn.executescript(
        """
        CREATE TABLE orders (
            id TEXT PRIMARY KEY,
            woo_id INTEGER,
            source TEXT,
            status TEXT,
            date_modified TEXT
        );
        CREATE TABLE order_notes (
            wc_note_id INTEGER,
            order_id TEXT,
            note TEXT,
            date_created TEXT,
            customer_note INTEGER,
            author TEXT,
            added_by_user INTEGER,
            UNIQUE(order_id, wc_note_id)
        );
        """
    )
    conn.executemany(
        """
        INSERT INTO orders (id, woo_id, source, status, date_modified)
        VALUES (?, ?, 'https://store.example', ?, ?)
        """,
        [
            ("1-101", 101, "processing", "2026-07-31T10:01:00"),
            ("1-102", 102, "on-hold", "2026-07-31T10:02:00"),
            ("1-103", 103, "processing", "2026-07-31T10:03:00"),
            ("1-104", 104, "completed", "2026-07-31T10:04:00"),
        ],
    )
    conn.commit()
    return conn


def test_subtract_minutes_preserves_timestamp_shape():
    assert (
        sync_utils._subtract_minutes_from_iso("2026-07-31T10:00:00", 10)
        == "2026-07-31T09:50:00"
    )
    assert (
        sync_utils._subtract_minutes_from_iso("2026-07-31T10:00:00Z", 10)
        == "2026-07-31T09:50:00Z"
    )
    assert (
        sync_utils._subtract_minutes_from_iso("2026-07-31T10:00:00+02:00", 10)
        == "2026-07-31T09:50:00+02:00"
    )


def test_note_refresh_includes_changed_orders_and_caps_active_rotation():
    conn = make_note_db()
    wcapi = FakeWCAPI()

    result = sync_utils.sync_order_notes(
        wcapi,
        "https://store.example",
        connection=conn,
        changed_woo_ids=[104],
        active_limit=2,
        max_workers=1,
    )

    assert result == {"candidates": 3, "synced": 3, "failed": 0, "notes": 3}
    assert [path for path, _ in wcapi.paths] == [
        "orders/104/notes",
        "orders/103/notes",
        "orders/102/notes",
    ]
    assert conn.execute("SELECT COUNT(*) FROM order_notes").fetchone()[0] == 3
    assert (
        conn.execute("SELECT COUNT(*) FROM order_note_sync_state").fetchone()[0]
        == 3
    )


def test_note_state_table_is_committed_when_no_orders_are_selected():
    conn = make_note_db()

    result = sync_utils.sync_order_notes(
        FakeWCAPI(),
        "https://missing.example",
        connection=conn,
        active_limit=0,
    )
    conn.close()

    assert result["candidates"] == 0


def test_sync_site_defaults_to_checkpoint_with_overlap(monkeypatch):
    conn = sqlite3.connect(":memory:")
    wcapi = object()
    modified_fetch = Mock(return_value=[{"id": 201}, {"id": 202}])
    full_fetch = Mock()
    note_sync = Mock(
        return_value={"candidates": 2, "synced": 2, "failed": 0, "notes": 0}
    )

    monkeypatch.setattr(sync_utils, "create_robust_wcapi", lambda *args: wcapi)
    monkeypatch.setattr(sync_utils, "get_thread_db_connection", lambda: conn)
    monkeypatch.setattr(sync_utils, "close_thread_db_connection", lambda: None)
    monkeypatch.setattr(
        sync_utils,
        "get_last_modified_date_from_db",
        lambda url: "2026-07-31T10:00:00",
    )
    monkeypatch.setattr(sync_utils, "fetch_orders_incrementally", full_fetch)
    monkeypatch.setattr(sync_utils, "fetch_orders_modified_after", modified_fetch)
    monkeypatch.setattr(sync_utils, "sync_order_notes", note_sync)

    result = sync_utils.sync_site(
        "https://store.example",
        "consumer-key",
        "consumer-secret",
    )

    full_fetch.assert_not_called()
    assert modified_fetch.call_args.args[2] == "2026-07-31T09:50:00"
    assert note_sync.call_args.kwargs["changed_woo_ids"] == [201, 202]
    assert note_sync.call_args.kwargs["active_limit"] == 25
    assert note_sync.call_args.kwargs["max_workers"] == 1
    assert result["updated_orders"] == 2
    assert result["notes_checked"] == 2
