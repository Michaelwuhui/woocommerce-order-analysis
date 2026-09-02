from pathlib import Path

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


def test_deep_entrypoint_submits_the_durable_pipeline():
    source = ROOT.joinpath("full_resync_all.py").read_text(encoding="utf-8")
    assert "from sync_service import start_sync" in source
    assert 'mode="deep"' in source
    assert 'params={"per_page": 50' in source
    assert "threading.Thread" not in source
    assert "woocommerce.API" not in source


def test_fetch_retry_budget_lives_in_the_celery_page_task():
    source = ROOT.joinpath("sync_tasks.py").read_text(encoding="utf-8")
    assert "MAX_FETCH_RETRIES = 3" in source
    assert "max_retries=MAX_FETCH_RETRIES" in source
    assert "raise self.retry(" in source


def test_fetch_retry_classification_is_explicit():
    source = ROOT.joinpath("sync_tasks.py").read_text(encoding="utf-8")
    assert "RETRYABLE_HTTP = {429, 500, 502, 503, 504}" in source
    assert "if status in {401, 403}:" in source
    assert "raise PermanentFetchError" in source


def test_runtime_progress_is_persisted_in_postgres_authority_tables():
    schema = ROOT.joinpath(
        "migrations", "postgresql", "002_sync_pipeline.sql"
    ).read_text(encoding="utf-8")
    for table in ("sync_runs", "sync_site_progress", "sync_page_receipts"):
        assert f"CREATE TABLE IF NOT EXISTS {table}" in schema
    for column in (
        "heartbeat_at", "completed_sites", "current_page", "fetched_count",
        "written_count", "retry_count",
    ):
        assert column in schema


def test_auto_and_full_sync_use_the_same_database_global_mutex():
    auto_source = ROOT.joinpath("auto_sync.py").read_text(encoding="utf-8")
    full_source = ROOT.joinpath("full_resync_all.py").read_text(encoding="utf-8")
    schema = ROOT.joinpath(
        "migrations", "postgresql", "002_sync_pipeline.sql"
    ).read_text(encoding="utf-8")

    assert "from sync_service import start_sync" in auto_source
    assert "from sync_service import start_sync" in full_source
    assert 'mode="auto"' in auto_source
    assert 'mode="deep"' in full_source
    assert "CREATE UNIQUE INDEX IF NOT EXISTS uq_sync_runs_one_active" in schema


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
