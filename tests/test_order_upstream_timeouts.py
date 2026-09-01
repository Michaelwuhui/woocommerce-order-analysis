import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_SOURCE = ROOT.joinpath("app.py").read_text(encoding="utf-8")
APP_TREE = ast.parse(APP_SOURCE)
FULFILLMENT_SOURCE = ROOT.joinpath("fulfillment_woocommerce.py").read_text(encoding="utf-8")
FULFILLMENT_TREE = ast.parse(FULFILLMENT_SOURCE)


def _function_source(name):
    node = next(
        item
        for item in APP_TREE.body
        if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)) and item.name == name
    )
    return ast.get_source_segment(APP_SOURCE, node)


def _fulfillment_function_source(name):
    node = next(
        item
        for item in FULFILLMENT_TREE.body
        if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)) and item.name == name
    )
    return ast.get_source_segment(FULFILLMENT_SOURCE, node)


def _constant(name):
    node = next(
        item
        for item in APP_TREE.body
        if isinstance(item, ast.Assign)
        and any(isinstance(target, ast.Name) and target.id == name for target in item.targets)
    )
    return ast.literal_eval(node.value)


def test_synchronous_woocommerce_timeouts_stay_bounded():
    assert _constant("WC_MUTATION_TIMEOUT") == (5, 15)
    assert _constant("WC_VERIFY_TIMEOUT") == (5, 10)
    assert _constant("WC_AUXILIARY_TIMEOUT") == (5, 10)


def test_shipping_mutation_is_sent_once_then_verified_without_blind_retry():
    source = _function_source("ship_order")

    assert source.count("request_method(") == 1
    assert "for attempt in range(3)" not in source
    assert "time.sleep(" not in source
    assert "timeout=WC_MUTATION_TIMEOUT" in source
    assert "timeout=WC_VERIFY_TIMEOUT" in source
    assert "except req.exceptions.RequestException" in source
    assert "remote_state_uncertain" in source
    assert "'retry_safe': not remote_state_uncertain" in source
    assert "请勿重复提交" in source


def test_shipping_2xx_response_must_contain_or_verify_exact_tracking():
    source = _function_source("ship_order")

    assert "if resp.status_code in (200, 201):\n            remote_success = True" not in source
    assert "if _remote_order_applied(resp.json()):" in source
    assert "成功响应未确认运单已写入" in source


def test_shipping_local_mirror_retries_without_repeating_remote_side_effects():
    source = _function_source("ship_order")

    assert "def _write_local_mirror(local_conn):" in source
    assert "_run_sqlite_write_with_retry(_write_local_mirror)" in source
    assert "'remote_committed': True" in source
    assert "'retry_safe': False" in source


def test_fulfillment_worker_commits_tracking_before_network_notification():
    source = _fulfillment_function_source("sync_shipment")
    local_mirror = source[source.index("exists = conn.execute("):]

    assert local_mirror.index("conn.commit()") < local_mirror.index("_notify_shipment(")
    notification_failure = local_mirror[local_mirror.index("except WooError:"):]
    assert notification_failure.index("conn.commit()") < notification_failure.index("raise")


def test_all_order_runtime_connections_wait_for_short_sqlite_contention():
    app_connection = _function_source("get_db_connection")
    auto_sync = ROOT.joinpath("auto_sync.py").read_text(encoding="utf-8")
    sync_utils = ROOT.joinpath("sync_utils.py").read_text(encoding="utf-8")

    assert "timeout_seconds=SQLITE_BUSY_TIMEOUT_SECONDS" in app_connection
    assert "timeout=timeout_seconds" in app_connection
    assert "PRAGMA busy_timeout" in app_connection
    assert "sqlite3.connect(DB_FILE, timeout=30)" in auto_sync
    assert "PRAGMA busy_timeout=30000" in auto_sync
    assert sync_utils.count("sqlite3.connect(DB_FILE, timeout=30)") >= 2


def test_status_mutation_verifies_ambiguous_response_before_local_commit():
    source = _function_source("update_order_status")

    assert source.count("req.put(") == 1
    assert "timeout=WC_MUTATION_TIMEOUT" in source
    assert "verify_resp = req.get(" in source
    assert "timeout=WC_VERIFY_TIMEOUT" in source
    assert source.index("if not remote_confirmed:") < source.index(
        'conn.execute("UPDATE orders SET status = ? WHERE id = ?"'
    )
    assert "'uncertain': remote_state_uncertain" in source
    assert "请勿重复提交" in source


def test_mutation_frontends_do_not_parse_proxy_html_as_json():
    base = ROOT.joinpath("templates", "base.html").read_text(encoding="utf-8")
    shipping = ROOT.joinpath("templates", "shipping.html").read_text(encoding="utf-8")
    settings = ROOT.joinpath("templates", "settings.html").read_text(encoding="utf-8")

    assert "function parseMutationJsonResponse(response, actionLabel)" in base
    assert "parseMutationJsonResponse(response, '订单状态')" in base
    assert "parseMutationJsonResponse(res, '发货')" in shipping
    assert "parseMutationJsonResponse(res, '补发')" in shipping
    assert "parseMutationJsonResponse(response, '订单同步')" in settings
    assert "parseMutationJsonResponse(res, '订单同步状态')" in settings
    assert "请勿重复提交" in base
    assert shipping.count("请求结果未知，请勿重复提交") >= 2


def test_web_service_supports_graceful_reload_for_in_flight_requests():
    drop_in = ROOT.joinpath(
        "deploy", "woo-analysis.service.d", "graceful-reload.conf"
    ).read_text(encoding="utf-8")

    assert "ExecReload=/bin/kill -HUP $MAINPID" in drop_in
    assert "--graceful-timeout 90" in drop_in


def test_manual_sync_status_is_published_for_all_gunicorn_workers():
    source = _function_source("sync_data")

    assert source.count("_publish_sync_status(site_id)") >= 6
    assert source.index("_publish_sync_status(site_id)") < source.index("thread.start()")
