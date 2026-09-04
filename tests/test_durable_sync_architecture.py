import ast
import inspect
import re
from pathlib import Path

import pytest
import requests

import sync_service
import sync_tasks
import db_backend
from celery_app import celery_app
from deploy.prepare_cutover_crontab import filtered_crontab


ROOT = Path(__file__).resolve().parents[1]


class _Response:
    def __init__(self, status_code, payload=None, headers=None):
        self.status_code = status_code
        self._payload = [] if payload is None else payload
        self.headers = headers or {}

    def json(self):
        return self._payload


class _Session:
    def __init__(self, response=None, error=None):
        self.response = response
        self.error = error

    def get(self, *_args, **_kwargs):
        if self.error:
            raise self.error
        return self.response


@pytest.mark.parametrize("status", [429, 500, 502, 503, 504, 599])
def test_only_retryable_http_classes_raise_transient(status):
    response = _Response(status, headers={"Retry-After": "7"})
    with pytest.raises(sync_tasks.TransientFetchError) as raised:
        sync_tasks._http_json_list(
            _Session(response), "https://unit.invalid", auth=("x", "y")
        )
    if status == 429:
        assert raised.value.retry_after == 7


@pytest.mark.parametrize("status", [401, 403])
def test_authentication_failures_are_never_blindly_retried(status):
    with pytest.raises(sync_tasks.AuthenticationFetchError):
        sync_tasks._http_json_list(
            _Session(_Response(status)), "https://unit.invalid", auth=("x", "y")
        )


@pytest.mark.parametrize("status", [400, 404, 409, 422])
def test_deterministic_client_errors_are_permanent(status):
    with pytest.raises(sync_tasks.PermanentFetchError):
        sync_tasks._http_json_list(
            _Session(_Response(status)), "https://unit.invalid", auth=("x", "y")
        )


@pytest.mark.parametrize(
    "error", [requests.Timeout("timeout"), requests.ConnectionError("down")]
)
def test_network_timeouts_and_disconnects_are_retryable(error):
    with pytest.raises(sync_tasks.TransientFetchError):
        sync_tasks._http_json_list(
            _Session(error=error), "https://unit.invalid", auth=("x", "y")
        )


def test_ipv4_preference_is_explicit_and_fetch_worker_scoped(monkeypatch):
    import socket
    from urllib3.util import connection as urllib3_connection

    original = urllib3_connection.allowed_gai_family
    try:
        monkeypatch.setenv("WOO_SYNC_IPV4_ONLY", "1")
        assert sync_tasks.configure_ipv4_preference() is True
        assert urllib3_connection.allowed_gai_family() == socket.AF_INET
    finally:
        urllib3_connection.allowed_gai_family = original


def test_fetch_payload_is_bounded_and_hash_verified():
    orders = [{"id": value} for value in range(100)]
    payload = {
        "run_id": "run",
        "site_id": 1,
        "page": 1,
        "orders": orders,
        "notes": [],
        "content_hash": sync_tasks._canonical_hash(orders, []),
    }
    assert sync_tasks._page_payload(payload)[3] == orders

    oversized = dict(payload, orders=orders + [{"id": 101}])
    oversized["content_hash"] = sync_tasks._canonical_hash(oversized["orders"], [])
    with pytest.raises(ValueError):
        sync_tasks._page_payload(oversized)

    corrupted = dict(payload, content_hash="0" * 64)
    with pytest.raises(ValueError):
        sync_tasks._page_payload(corrupted)


def test_psycopg_percent_escape_preserves_only_generated_placeholders():
    compiled = "SELECT sqlite_strftime('%Y-%m', %s), %s LIKE 'pytest:%'"
    assert db_backend._escape_psycopg_percents(compiled) == (
        "SELECT sqlite_strftime('%%Y-%%m', %s), %s LIKE 'pytest:%%'"
    )


def test_compat_row_iterates_values_and_still_converts_to_mapping():
    row = db_backend.CompatRow(["alpha", "beta"], [1, 2])
    assert list(row) == [1, 2]
    assert dict(row) == {"alpha": 1, "beta": 2}


def test_explicit_begin_bypasses_postgres_sql_compilation():
    class Info:
        transaction_status = db_backend.TransactionStatus.IDLE

    class RawConnection:
        info = Info()

    class RawCursor:
        def __init__(self):
            self.statements = []

        def execute(self, statement):
            self.statements.append(statement)

    class Connection:
        def __init__(self):
            self._raw = RawConnection()
            self._last_changes = 9

        def compile(self, *_args, **_kwargs):
            raise AssertionError("BEGIN must not invoke SQL compilation")

    connection = Connection()
    raw_cursor = RawCursor()
    cursor = db_backend.PgCompatCursor(connection, raw_cursor)

    cursor.execute("BEGIN")

    assert raw_cursor.statements == []
    assert connection._last_changes == 0


def test_replace_conflict_keys_match_the_migrated_reconciliation_schema():
    assert db_backend._REPLACE_CONFLICT_KEYS["reconciliation_statements"] == (
        "partner_id",
        "period_year",
        "period_month",
    )


def test_page_size_is_clamped_to_required_range():
    assert sync_service._bounded_per_page(1) == 50
    assert sync_service._bounded_per_page(75) == 75
    assert sync_service._bounded_per_page(1000) == 100


def test_next_same_site_page_is_dispatched_only_after_receipt_write():
    source = inspect.getsource(sync_tasks._write_page_transaction)
    assert source.index("INSERT INTO sync_page_receipts") < source.index(
        "enqueue_fetch_page(connection, run_id, site_id, page + 1)"
    )
    fetch_source = inspect.getsource(sync_tasks.fetch_page)
    assert "_site_lock(lock_connection, site_id)" in fetch_source
    assert "note_retry" not in inspect.getsource(sync_tasks._defer_busy_site)


def test_celery_is_json_only_late_ack_and_prefetch_one():
    assert celery_app.conf.task_serializer == "json"
    assert celery_app.conf.accept_content == ["json"]
    assert celery_app.conf.result_serializer == "json"
    assert celery_app.conf.worker_prefetch_multiplier == 1
    assert celery_app.conf.task_acks_late is True
    assert celery_app.conf.task_reject_on_worker_lost is True
    assert celery_app.conf.broker_transport_options["visibility_timeout"] >= 1800
    assert "pickle" not in repr(dict(celery_app.conf)).lower()
    assert celery_app.conf.task_routes["woo_sync.post_commit_page"]["queue"] == "sync_write"


def test_post_commit_hooks_have_postgres_receipt_and_stale_recovery():
    migration = (ROOT / "migrations/postgresql/003_sync_post_commit.sql").read_text()
    transaction = inspect.getsource(sync_tasks._write_page_transaction)
    recovery = inspect.getsource(sync_service.recover_stale_work)
    task = inspect.getsource(sync_tasks.post_commit_page)
    assert "planning_candidates jsonb" in migration
    assert "post_commit_status" in migration
    assert 'task_name="woo_sync.post_commit_page"' in transaction
    assert "postcommit:" in recovery
    assert "strict=True" in task


def test_worker_units_enforce_fetch_three_writer_one_and_single_beat():
    fetch = (ROOT / "deploy/systemd/woo-celery-fetch.service").read_text()
    writer = (ROOT / "deploy/systemd/woo-celery-write.service").read_text()
    beat = (ROOT / "deploy/systemd/woo-celery-beat.service").read_text()
    assert "--queues=sync_fetch" in fetch
    assert "--concurrency=3" in fetch
    assert "Environment=WOO_DB_POOL_MAX=2" in fetch
    assert "Environment=WOO_SYNC_IPV4_ONLY=1" in fetch
    assert "--queues=sync_write" in writer
    assert "--concurrency=1" in writer
    # Foreground workers are supervised by systemd and must not share Beat's
    # RuntimeDirectory. Restarting either worker used to remove the other
    # process's pidfile and make its next clean shutdown fail with EROFS.
    assert "--pidfile=" not in fetch
    assert "--pidfile=" not in writer
    assert "RuntimeDirectory=woo-analysis" not in fetch
    assert "RuntimeDirectory=woo-analysis" not in writer
    assert "--pidfile=/run/woo-analysis/celery-beat.pid" in beat
    assert "Restart=always" in fetch
    assert "Restart=always" in writer
    assert "Restart=always" in beat


def test_beat_checks_database_backed_auto_and_deep_schedules():
    schedules = celery_app.conf.beat_schedule
    assert schedules["automatic-sync-due-check"]["schedule"] == 60.0
    assert schedules["deep-sync-due-check"]["schedule"] == 60.0
    auto_source = inspect.getsource(sync_tasks.schedule_auto)
    deep_source = inspect.getsource(sync_tasks.schedule_deep)
    assert "_auto_due()" in auto_source
    assert "_deep_due()" in deep_source
    assert 'created_by="celery-beat:auto"' in auto_source
    assert 'created_by="celery-beat:deep"' in deep_source
    migration = (ROOT / "migrations/postgresql/002_sync_pipeline.sql").read_text()
    assert "INSERT INTO settings" not in migration


def test_postgresql_mode_cannot_recreate_legacy_sync_crons():
    source = (ROOT / "app.py").read_text()
    autosync = source[source.index("def set_autosync_status") : source.index("@app.route('/api/sync/logs'")]
    deep = source[source.index("def setup_cron") : source.index("def remove_cron")]
    clean = source[source.index("def setup_clean_cron") : source.index("def remove_clean_cron")]
    for function_source in (autosync, deep, clean):
        assert "sqlite3.is_postgres_backend()" in function_source
    assert autosync.index("sqlite3.is_postgres_backend()") < autosync.index("/usr/bin/crontab")
    assert deep.index("sqlite3.is_postgres_backend()") < deep.index("/usr/bin/crontab")
    assert clean.index("sqlite3.is_postgres_backend()") < clean.index("/usr/bin/crontab")


def test_products_route_ends_read_transactions_before_cpu_aggregation():
    source = (ROOT / "app.py").read_text()
    products = source[source.index("def products():") : source.index("@app.route('/api/brands')")]
    first_query = products.index("mappings_rows = conn.execute")
    first_commit = products.index("conn.commit()", first_query)
    assert first_commit < products.index("# Aggregate product data")

    loss_query = products.index("shipping_loss_by_currency = {}")
    loss_commit = products.index("conn.commit()", loss_query)
    assert loss_commit < products.index("# Subtract the loss")

    trend_query = products.index("trend_orders = conn.execute")
    trend_commit = products.index("conn.commit()", trend_query)
    assert trend_commit < products.index("weekly_flavor_data = {}")


def test_cutover_crontab_filter_is_exact_and_preserves_unrelated_jobs():
    source = """LANG=en_US.UTF-8
30 3 * * * cd /www/wwwroot/woo-analysis && python 1.wooorders_sqlite.py
30 4 * * 0 cd /www/wwwroot/woo-analysis && python 1.wooorders_sqlite.py --clean
17 * * * * /usr/bin/flock -n /run/lock/woo-analysis-auto-sync.cron.lock python auto_sync.py
20 * * * * python /www/wwwroot/woo-analysis/backup_db.py
*/5 * * * * cd /www/wwwroot/woo-analysis && /usr/bin/flock -n /run/lock/woo-analysis-inv-push.cron.lock python inv_push_cron.py
0 5 * * * python /www/wwwroot/woo-analysis/resolve_outcomes.py --live
"""
    maintenance, counts = filtered_crontab(source, "maintenance")
    assert counts == {
        "deep_and_clean": 2,
        "automatic_sync": 1,
        "sqlite_backup": 1,
        "inventory_push": 1,
        "resolve_outcomes": 1,
    }
    for command in (
        "1.wooorders_sqlite.py",
        "auto_sync.py",
        "backup_db.py",
        "inv_push_cron.py",
        "resolve_outcomes.py",
    ):
        assert command not in maintenance
    assert maintenance == "LANG=en_US.UTF-8\n"

    postgres, postgres_counts = filtered_crontab(source, "postgres")
    assert postgres_counts == counts
    assert "1.wooorders_sqlite.py" not in postgres
    assert "auto_sync.py" not in postgres
    assert "backup_db.py" not in postgres
    assert postgres.count(". /etc/woo-analysis/woo-analysis.env") == 2
    assert "inv_push_cron.py" in postgres
    assert "resolve_outcomes.py" in postgres
    assert postgres.startswith("LANG=en_US.UTF-8\n")


def test_cutover_crontab_filter_refuses_changed_legacy_counts():
    with pytest.raises(ValueError, match="managed cron counts changed"):
        filtered_crontab("17 * * * * python auto_sync.py\n", "maintenance")


def test_non_celery_workers_use_postgres_skip_locked_claims():
    fulfillment = (ROOT / "fulfillment_worker.py").read_text()
    product_clone = (ROOT / "product_clone_jobs.py").read_text()
    assert "FOR UPDATE SKIP LOCKED" in fulfillment
    assert "FOR UPDATE SKIP LOCKED" in product_clone


def test_redis_is_local_durable_and_noeviction():
    config = (ROOT / "deploy/redis/woo-analysis.conf").read_text()
    sysctl = (ROOT / "deploy/sysctl/99-woo-analysis-redis.conf").read_text()
    service_dropin = (
        ROOT / "deploy/systemd/redis-server.service.d/woo-analysis.conf"
    ).read_text()
    assert "bind 127.0.0.1" in config
    assert "protected-mode yes" in config
    assert "appendonly yes" in config
    assert "appendfsync everysec" in config
    assert any(line.startswith("save ") for line in config.splitlines())
    assert "maxmemory-policy noeviction" in config
    assert "vm.overcommit_memory = 1" in sysctl
    assert "LD_LIBRARY_PATH=/lib/x86_64-linux-gnu" in service_dropin
    assert "ExecPaths=/usr/local/lib" not in service_dropin


def test_backups_stage_atomic_artifacts_on_destination_filesystem():
    source = (ROOT / "backup_db.py").read_text()
    assert source.count('tempfile.mkdtemp(prefix=".woo-') == 2
    assert source.count('dir=BACKUP_DIR') == 2


def test_backup_supports_both_backends_checksums_and_hourly_timer():
    backup = (ROOT / "backup_db.py").read_text()
    service = (ROOT / "deploy/systemd/woo-postgres-backup.service").read_text()
    timer = (ROOT / "deploy/systemd/woo-postgres-backup.timer").read_text()
    assert "source.backup(destination)" in backup
    assert '"PRAGMA integrity_check"' in backup
    assert '"--format=custom"' in backup
    assert 'pg_restore, "--list"' in backup
    assert "sha256_file" in backup
    assert "EnvironmentFile=/etc/woo-analysis/woo-analysis.env" in service
    assert "OnCalendar=hourly" in timer


def test_global_sync_has_one_frontend_binding_and_one_post_site():
    base = (ROOT / "templates/base.html").read_text()
    settings = (ROOT / "templates/settings.html").read_text()
    runtime = (ROOT / "static/js/sync_runs.js").read_text()
    combined = base + settings + runtime
    assert combined.count("fetch('/api/sync/all'") == 1
    assert combined.count("button.addEventListener('click', startGlobalSync)") == 1
    assert "syncAllBtn.addEventListener" not in base
    assert "syncAllBtn.addEventListener" not in settings
    assert "任务已中断/正在恢复" in runtime
    for field in (
        "syncModeValue", "syncSiteValue", "syncSitesValue", "syncPageValue",
        "syncFetchedValue", "syncWrittenValue", "syncRetryValue",
        "syncHeartbeatValue",
    ):
        assert field in base
        assert field in runtime


def test_single_site_sync_polls_the_durable_run_id_and_handles_all_terminals():
    settings = (ROOT / "templates/settings.html").read_text()
    restricted = (ROOT / "static/js/site_sync_settings.js").read_text()

    assert "const syncId = result.run_id || result.sync_id || siteId" in settings
    assert "sync/status/${encodeURIComponent(syncId)}" in settings
    assert "sync/status/${siteId}" not in settings
    assert "['error', 'cancelled', 'interrupted']" in settings
    assert "任务已中断/正在恢复" in settings

    assert "data.run_id || data.sync_id || siteId" in restricted
    for status in ("success", "error", "cancelled", "interrupted"):
        assert status in restricted


def test_postgres_latest_note_queries_and_request_cleanup_are_deterministic():
    source = (ROOT / "app.py").read_text()
    assert "GROUP BY order_id HAVING date_created = MAX(date_created)" not in source
    assert "GROUP BY order_id\n            HAVING date_created = MAX(date_created)" not in source
    assert "ORDER BY n2.date_created DESC, n2.id DESC LIMIT 1" in source
    assert "@app.teardown_request" in source
    assert "_close_request_postgres_connections" in source


def test_analytics_routes_reuse_connections_and_customer_fields_are_aggregated():
    source = (ROOT / "app.py").read_text()
    duplicate_open = (
        "# Get available sources (filtered by permissions and manager)\n"
        "    conn = get_db_connection()"
    )
    assert duplicate_open not in source
    assert "MAX(json_extract(billing, '$.first_name'))" in source
    assert "MAX(json_extract(billing, '$.last_name'))" in source
    assert "MAX(json_extract(billing, '$.phone')) as phone" in source
    assert "CASE WHEN json_valid(orders.line_items) THEN" not in source
    assert "CASE WHEN json_valid(orders.line_items) = 1 THEN" in source
    assert "CAST(COALESCE(json_extract(item.value, '$.quantity'), 0) AS INTEGER)" not in source
    assert "COALESCE(CAST(json_extract(item.value, '$.quantity') AS INTEGER), 0)" in source


def test_no_fixed_global_sync_status_id_remains_in_runtime_code():
    python_sources = [
        ROOT / "app.py",
        ROOT / "auto_sync.py",
        ROOT / "full_resync_all.py",
    ]
    for path in python_sources:
        tree = ast.parse(path.read_text())
        assert not any(
            isinstance(node, ast.Constant)
            and isinstance(node.value, int)
            and node.value == 999999
            for node in ast.walk(tree)
        )
    assert "999999" not in (ROOT / "static/js/sync_runs.js").read_text()


def test_bulk_api_check_has_no_live_write_probe_or_postgres_web_thread():
    source = (ROOT / "app.py").read_text()
    route = source[
        source.index("def check_all_sites_api") :
        source.index("def get_check_status")
    ]
    assert "sqlite3.is_postgres_backend()" in route
    assert route.index("sqlite3.is_postgres_backend()") < route.index(
        "threading.Thread"
    )
    assert "new_sync_runtime_status_id()" in route
    assert "888888" not in route
    assert "wcapi.post" not in route
    assert "wcapi.delete" not in route


def test_postgres_guards_every_legacy_gunicorn_background_thread():
    source = (ROOT / "app.py").read_text()
    tree = ast.parse(source)
    threaded_functions = []
    for node in tree.body:
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        function_source = ast.get_source_segment(source, node) or ""
        if "threading.Thread" not in function_source:
            continue
        threaded_functions.append(node.name)
        assert "sqlite3.is_postgres_backend()" in function_source
        assert function_source.index("sqlite3.is_postgres_backend()") < (
            function_source.index("threading.Thread")
        )
    assert threaded_functions == [
        "sync_data",
        "deep_sync_site",
        "check_all_sites_api",
        "clean_sync_site",
        "sync_all_data",
        "trigger_deep_sync",
        "clean_all_sites",
    ]


def test_root_diagnostics_are_read_only():
    for name in ("test_api_check.py", "diagnose_api.py"):
        source = (ROOT / name).read_text()
        assert "wcapi.post" not in source
        assert "wcapi.put" not in source
        assert "wcapi.delete" not in source


def test_legacy_scripts_have_no_source_controlled_woocommerce_credentials():
    credential = re.compile(r"['\"](?:ck|cs)_[A-Za-z0-9]{20,}['\"]")
    for name in ("1.wooorders_sqlite.py", "debug_cleanup_issue.py"):
        source = (ROOT / name).read_text()
        assert credential.search(source) is None
    legacy = (ROOT / "1.wooorders_sqlite.py").read_text()
    assert "HARDCODED_SITES" not in legacy
    assert "WOO_SQLITE_PATH" in legacy


def test_pytest_collection_excludes_live_root_diagnostic_scripts():
    config = (ROOT / "pytest.ini").read_text()
    assert "testpaths = tests" in config
