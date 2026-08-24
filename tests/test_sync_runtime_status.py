import sqlite3

from sync_runtime_status import (
    init_sync_runtime_status,
    load_sync_runtime_status,
    new_sync_runtime_status_id,
    save_sync_runtime_status,
)


def connect(path):
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


def test_generated_status_ids_are_distinct_and_browser_safe():
    status_ids = [new_sync_runtime_status_id() for _ in range(1000)]

    assert len(set(status_ids)) == len(status_ids)
    assert all(0 < status_id <= (2**53 - 1) for status_id in status_ids)
    assert 999999 not in status_ids


def test_status_is_shared_across_connections(tmp_path):
    db_path = tmp_path / "shared-sync-status.db"
    writer = connect(db_path)
    reader = connect(db_path)
    init_sync_runtime_status(writer)

    save_sync_runtime_status(
        writer,
        999999,
        {
            "status": "running",
            "message": "Syncing site 2/32",
            "logs": ["started", "site 1 completed"],
            "updated_at": 1000.0,
        },
    )
    loaded = load_sync_runtime_status(reader, 999999, now=1004.25)
    assert loaded["status"] == "running"
    assert loaded["message"] == "Syncing site 2/32"
    assert loaded["logs"] == ["started", "site 1 completed"]
    assert loaded["stale_seconds"] == 4.2
    writer.close()
    reader.close()


def test_new_run_overwrites_terminal_status_for_other_workers(tmp_path):
    db_path = tmp_path / "replace-sync-status.db"
    conn = connect(db_path)
    init_sync_runtime_status(conn)
    save_sync_runtime_status(
        conn,
        999999,
        {"status": "success", "message": "old run", "logs": [], "updated_at": 10},
    )
    save_sync_runtime_status(
        conn,
        999999,
        {"status": "running", "message": "new run", "logs": ["new"], "updated_at": 20},
    )
    loaded = load_sync_runtime_status(conn, 999999, now=21)
    assert loaded["status"] == "running"
    assert loaded["message"] == "new run"
    assert loaded["logs"] == ["new"]
    assert loaded["stale_seconds"] == 1.0
    conn.close()
