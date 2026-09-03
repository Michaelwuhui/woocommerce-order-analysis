import importlib.util
from pathlib import Path

import pytest

import inv_migrations
from inv_skus import _parse_suggest_candidate_query


ROOT = Path(__file__).resolve().parents[1]


def test_optional_sku_filters_never_bind_untyped_null_predicates():
    query, params = _parse_suggest_candidate_query(7, None, None)
    assert params == (7,)
    assert '? IS NULL' not in query
    assert 'series_id=?' not in query
    assert 'puff_count=?' not in query

    query, params = _parse_suggest_candidate_query(7, 11, 40000)
    assert params == (7, 11, 40000)
    assert '(series_id=? OR series_id IS NULL)' in query
    assert '(puff_count=? OR puff_count IS NULL)' in query


def test_postgres_route_comparison_is_null_safe_and_typed():
    source = (ROOT / 'order_notification_api.py').read_text(encoding='utf-8')
    assert 'store_id IS NOT DISTINCT FROM ?' in source
    assert 'warehouse_id IS NOT DISTINCT FROM ?' in source
    assert 'store_id IS ?' not in source
    assert 'warehouse_id IS ?' not in source


def test_legacy_inventory_migrator_refuses_postgres_before_connect(monkeypatch):
    monkeypatch.setenv('WOO_DB_BACKEND', 'postgres')
    monkeypatch.setattr(
        inv_migrations,
        'get_conn',
        lambda: pytest.fail('PostgreSQL guard must run before opening a connection'),
    )

    with pytest.raises(SystemExit, match='仅支持 SQLite'):
        inv_migrations.cmd_up()
    with pytest.raises(SystemExit, match='仅支持 SQLite'):
        inv_migrations.cmd_down()


def test_legacy_sqlite_sync_refuses_postgres_before_any_remote_work(monkeypatch):
    module_path = ROOT / '1.wooorders_sqlite.py'
    spec = importlib.util.spec_from_file_location('legacy_woo_sqlite_sync', module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    monkeypatch.setenv('WOO_DB_BACKEND', 'postgres')

    with pytest.raises(SystemExit, match='旧 SQLite 同步器'):
        module.main()


def test_known_postgres_incompatible_fragments_are_absent_from_runtime_paths():
    runtime_sources = [
        ROOT / 'app.py',
        ROOT / 'sync_utils.py',
        ROOT / 'order_notification_api.py',
        ROOT / 'inv_skus.py',
    ]
    combined = '\n'.join(path.read_text(encoding='utf-8') for path in runtime_sources)
    assert "strftime('%Y-W%W'" not in combined
    assert 'order_notes.added_by_user = 1 THEN 1' not in combined
    assert '? IS NULL OR series_id=?' not in combined
    assert '? IS NULL OR puff_count=?' not in combined


def test_company_profit_period_lookup_avoids_sqlite_printf():
    source = (ROOT / 'company_profit.py').read_text(encoding='utf-8')
    assert "printf('%04d-%02d'" not in source
    assert '(period_year * 100 + period_month) < ?' in source
