"""Real routes/authentication/transactions against an EMPTY production-schema copy."""
import os
import json
from concurrent.futures import ThreadPoolExecutor

import pytest
import requests

import db_backend as db
from return_shipping_loss import SETTING_KEY, PolicyConflict, load_policy, save_policy


pytestmark = pytest.mark.skipif(not db.is_postgres_backend(), reason='isolated PostgreSQL required')
OID = '99018-1'
RULES = [{'warehouse_id': None, 'destination_country': 'CZ', 'currency': 'CZK',
          'outbound_amount': 200, 'return_amount': 200}]


@pytest.fixture
def api(monkeypatch):
    assert os.environ.get('WOO_DB_NAME_OVERRIDE', '').startswith('woo_return_loss_test_')
    monkeypatch.setattr(requests.sessions.Session, 'request',
                        lambda *a, **k: pytest.fail('Real network calls forbidden'))
    import app as module
    module.app.config.update(TESTING=True)
    c = db.connect()
    c.execute('DELETE FROM settings WHERE key=?', (SETTING_KEY,))
    c.execute('DELETE FROM oms_shipments WHERE fulfillment_id IN (SELECT id FROM oms_fulfillments WHERE order_id=?)', (OID,))
    c.execute('DELETE FROM oms_fulfillments WHERE order_id=?', (OID,))
    c.execute('DELETE FROM oms_order_fulfillment_state WHERE order_id=?', (OID,))
    c.execute('DELETE FROM order_notes WHERE order_id=?', (OID,))
    c.execute('DELETE FROM orders WHERE id=?', (OID,))
    c.execute("INSERT INTO warehouses(id,name,code,country,default_currency) "
              "VALUES (99018,'Return test','RETURN-TEST','PL','PLN') ON CONFLICT(id) DO NOTHING")
    c.execute("INSERT INTO sites(id,url,consumer_key,consumer_secret,country) "
              "VALUES (99018,'https://return-test.invalid','','','CZ') ON CONFLICT(id) DO NOTHING")
    for uid, username, role, manage, ship in [
        (99018, 'return-test-admin', 'admin', True, True),
        (99019, 'return-test-user', 'user', False, False),
        (99020, 'return-test-shipper', 'user', False, True),
    ]:
        c.execute('''INSERT INTO users(id,username,password_hash,name,role,can_manage_users,can_ship)
                     VALUES (?,?,'unusable','Return test',?,?,?) ON CONFLICT(id) DO NOTHING''',
                  (uid, username, role, manage, ship))
    c.execute('''INSERT INTO orders(id,number,source,status,currency,shipping_total,shipping,billing,
                 line_items,meta_data,shipping_lines,date_created,payment_method)
                 VALUES (?,'1','https://return-test.invalid','shipped','CZK',0,?,'{}','[]','[]','[]',
                         '2026-09-01T00:00:00','cod')''', (OID, json.dumps({'country': 'CZ'})))
    c.commit()
    c.close()

    def client(uid=99018):
        client = module.app.test_client()
        if uid:
            with client.session_transaction() as session:
                session['_user_id'] = str(uid)
                session['_fresh'] = True
        return client
    return client


def save(client, rules=RULES):
    old = client.get('/api/settings/return-shipping-loss')
    assert old.status_code == 200
    return client.post('/api/settings/return-shipping-loss',
                       json={'rules': rules, 'version': old.json['version']})


def test_settings_persist_and_render_for_authorized_user(api):
    client = api()
    response = save(client)
    assert response.status_code == 200 and response.json['rules'] == RULES
    assert client.get('/api/settings/return-shipping-loss').json['rules'] == RULES
    page = client.get('/settings')
    assert page.status_code == 200
    assert '退件运费损失'.encode() in page.data
    assert client.get(f'/api/order/{OID}/return-shipping-loss').json['amount'] == 400


def test_settings_denied_without_global_settings_permission(api):
    for uid in (99019, 99020):
        client = api(uid)
        assert client.get('/api/settings/return-shipping-loss').status_code == 403
        assert client.post('/api/settings/return-shipping-loss', json={'rules': []}).status_code == 403
    assert api(None).get('/api/settings/return-shipping-loss').status_code == 302
    assert api(99019).get(f'/api/order/{OID}/return-shipping-loss').status_code == 403


def test_manual_return_uses_policy_and_retains_audit(api):
    client = api()
    assert save(client).status_code == 200
    response = client.post(f'/api/order/{OID}/mark-undelivered', json={})
    assert response.status_code == 200
    c = db.connect()
    order = c.execute('SELECT * FROM orders WHERE id=?', (OID,)).fetchone()
    assert order['shipping_loss_amount'] == 400 and order['is_undelivered']
    note = c.execute('SELECT note FROM order_notes WHERE order_id=?', (OID,)).fetchone()[0]
    assert '400.00' in note
    c.close()
    assert client.post(f'/api/order/{OID}/mark-undelivered', json={}).status_code == 409


@pytest.mark.parametrize('amount', [0, 375.5])
def test_manual_actual_amount_is_retained_when_adding_problem_return(api, amount):
    client = api()
    assert save(client).status_code == 200
    assert client.post(f'/api/order/{OID}/mark-undelivered', json={'shipping_loss_amount': amount}).status_code == 200
    assert client.get(f'/api/order/{OID}/return-shipping-loss').json['amount'] == amount
    result = client.post(f'/api/order/{OID}/mark-problem-return', json={'type': 'swap', 'product_loss_amount': 100})
    assert result.status_code == 200
    c = db.connect()
    assert c.execute('SELECT shipping_loss_amount FROM orders WHERE id=?', (OID,)).fetchone()[0] == amount
    c.close()


@pytest.mark.parametrize('amount', ['NaN', 'Infinity', -1, None, ''])
def test_invalid_manual_amount_does_not_mark_order(api, amount):
    result = api().post(f'/api/order/{OID}/mark-undelivered', json={'shipping_loss_amount': amount})
    assert result.status_code == 400
    c = db.connect()
    assert not c.execute('SELECT is_undelivered FROM orders WHERE id=?', (OID,)).fetchone()[0]
    c.close()


def test_problem_return_alone_uses_round_trip_loss(api):
    client = api()
    assert save(client).status_code == 200
    result = client.post(f'/api/order/{OID}/mark-problem-return', json={'type': 'swap', 'product_loss_amount': 100})
    assert result.status_code == 200
    c = db.connect()
    assert c.execute('SELECT shipping_loss_amount FROM orders WHERE id=?', (OID,)).fetchone()[0] == 400
    c.close()


def test_actual_warehouse_rule_on_postgresql(api):
    client = api()
    assert save(client, RULES + [dict(RULES[0], warehouse_id=99018, outbound_amount=150)]).status_code == 200
    c = db.connect()
    c.execute("INSERT INTO oms_order_fulfillment_state(order_id,revision,aggregate_status) VALUES (?,1,'shipped')", (OID,))
    c.execute('''INSERT INTO oms_fulfillments(id,order_id,warehouse_id,revision,status,idempotency_key,plan_hash)
                 VALUES ('return-test-f',?,99018,1,'shipped','return-test-f','test')''', (OID,))
    c.execute("INSERT INTO oms_shipments(id,fulfillment_id,tracking_number,status) VALUES ('return-test-s','return-test-f','TEST','returned')")
    c.commit()
    c.close()
    assert client.get(f'/api/order/{OID}/return-shipping-loss').json['amount'] == 350


def test_concurrent_first_saves_have_exactly_one_winner(api):
    c = db.connect()
    version = load_policy(c)['version']
    c.close()
    def attempt(amount):
        connection = db.connect()
        try:
            save_policy(connection, [dict(RULES[0], outbound_amount=amount)], version)
            connection.commit()
            return 'saved'
        except PolicyConflict:
            connection.rollback()
            return 'conflict'
        finally:
            connection.close()
    with ThreadPoolExecutor(max_workers=2) as pool:
        assert sorted(pool.map(attempt, [100, 300])) == ['conflict', 'saved']


def test_stale_http_save_rejected(api):
    client = api()
    version = client.get('/api/settings/return-shipping-loss').json['version']
    assert save(client).status_code == 200
    result = client.post('/api/settings/return-shipping-loss', json={'rules': [], 'version': version})
    assert result.status_code == 409
    assert client.get('/api/settings/return-shipping-loss').json['rules'] == RULES
