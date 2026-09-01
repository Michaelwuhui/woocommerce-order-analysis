from pathlib import Path

import pytest

import full_resync_all
from sync_runtime_status import load_sync_runtime_status


ROOT = Path(__file__).resolve().parents[1]


class FakeResponse:
    def __init__(self, status_code=200, payload=None):
        self.status_code = status_code
        self._payload = [] if payload is None else payload
        self.text = "temporary upstream failure"
        self.headers = {}

    def json(self):
        return self._payload


class SequenceWCAPI:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.calls = 0

    def get(self, path, params=None):
        self.calls += 1
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


def test_page_request_retries_network_failures_then_succeeds():
    wcapi = SequenceWCAPI(
        [TimeoutError("ipv6 timeout"), TimeoutError("read timeout"), FakeResponse()]
    )
    sleeps = []

    response = full_resync_all.fetch_orders_page(
        wcapi,
        7,
        sleep_fn=sleeps.append,
    )

    assert response.status_code == 200
    assert wcapi.calls == full_resync_all.MAX_PAGE_ATTEMPTS == 3
    assert sleeps == [1, 2]


def test_page_request_stops_after_bounded_attempts():
    wcapi = SequenceWCAPI([TimeoutError("offline")] * 5)

    with pytest.raises(TimeoutError, match="offline"):
        full_resync_all.fetch_orders_page(wcapi, 3, sleep_fn=lambda _seconds: None)

    assert wcapi.calls == full_resync_all.MAX_PAGE_ATTEMPTS == 3


def test_page_request_retries_retryable_http_status():
    wcapi = SequenceWCAPI([FakeResponse(503), FakeResponse(200)])

    response = full_resync_all.fetch_orders_page(
        wcapi,
        2,
        sleep_fn=lambda _seconds: None,
    )

    assert response.status_code == 200
    assert wcapi.calls == 2


def test_runtime_reporter_persists_progress_for_other_processes(tmp_path, monkeypatch):
    db_path = tmp_path / "orders.db"
    monkeypatch.setattr(full_resync_all, "DB_FILE", str(db_path))
    reporter = full_resync_all.RuntimeStatusReporter(123456)

    reporter.update(
        status="running",
        message="站点 2/34，第 4 页",
        progress=4.25,
        log="page 4 saved",
    )

    conn = full_resync_all._connect_db()
    loaded = load_sync_runtime_status(conn, 123456)
    conn.close()
    assert loaded["status"] == "running"
    assert loaded["message"] == "站点 2/34，第 4 页"
    assert loaded["progress"] == 4.25
    assert loaded["logs"] == ["page 4 saved"]


def test_auto_and_full_sync_use_the_same_exclusive_lock():
    auto_source = ROOT.joinpath("auto_sync.py").read_text(encoding="utf-8")
    full_source = ROOT.joinpath("full_resync_all.py").read_text(encoding="utf-8")

    assert "from sync_process_lock import" in auto_source
    assert "from sync_process_lock import" in full_source
    assert "with exclusive_sync_lock()" in auto_source
    assert "with exclusive_sync_lock()" in full_source


def test_deep_sync_route_uses_unique_persisted_status_and_child_status_id():
    source = ROOT.joinpath("app.py").read_text(encoding="utf-8")
    route = source[source.index("def trigger_deep_sync():"):source.index(
        "def clean_all_sites():"
    )]

    assert "new_sync_runtime_status_id()" in route
    assert "_publish_sync_status(deep_sync_id)" in route
    assert "'--status-id', str(deep_sync_id)" in route
    assert "DEEP_SYNC_ID = 888888" not in route


def test_deep_sync_ui_uses_persisted_numeric_progress():
    source = ROOT.joinpath("templates", "settings.html").read_text(encoding="utf-8")

    assert "status.progress" in source
