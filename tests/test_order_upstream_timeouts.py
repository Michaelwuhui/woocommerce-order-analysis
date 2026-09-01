import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_SOURCE = ROOT.joinpath("app.py").read_text(encoding="utf-8")
APP_TREE = ast.parse(APP_SOURCE)


def _function_source(name):
    node = next(
        item
        for item in APP_TREE.body
        if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)) and item.name == name
    )
    return ast.get_source_segment(APP_SOURCE, node)


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


def test_manual_sync_status_is_published_for_all_gunicorn_workers():
    source = _function_source("sync_data")

    assert source.count("_publish_sync_status(site_id)") >= 6
    assert source.index("_publish_sync_status(site_id)") < source.index("thread.start()")
