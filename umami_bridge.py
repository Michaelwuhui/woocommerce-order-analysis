#!/usr/bin/env python3
"""HK 订单系统与 CZ1 Umami 总览服务之间的源站 HTTPS 安全桥接客户端。"""
import datetime as dt
import os
import pathlib
import re
import ssl

import requests
from requests.adapters import HTTPAdapter
from urllib3.poolmanager import PoolManager

ENV_PATH = pathlib.Path('/etc/woo-analysis/umami-sync.env')

COUNTRY_TO_REGION = {
    'PL': '波兰',
    'CZ': '捷克',
    'HU': '匈牙利',
    'AU': '澳大利亚',
}
REGION_TO_COUNTRY = {
    '波兰': 'PL',
    '捷克': 'CZ',
    '匈牙利': 'HU',
    '澳大利亚': 'AU',
    '其他': 'OTHER',
}
UUID_RE = re.compile(r'^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$', re.I)


def _load_env():
    if not os.environ.get('UMAMI_BRIDGE_TOKEN') and ENV_PATH.is_file():
        for line in ENV_PATH.read_text(encoding='utf-8').splitlines():
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def normalize_domain(url: str) -> str:
    if not url:
        return ''
    u = str(url).strip().lower()
    u = re.sub(r'^https?://', '', u)
    u = u.split('/')[0].split('?')[0].split('#')[0].split(':')[0].rstrip('.')
    if u.startswith('www.'):
        u = u[4:]
    return u


def country_to_region(code: str) -> str:
    c = str(code or '').strip().upper()
    if not c:
        return ''
    return COUNTRY_TO_REGION.get(c, '其他')


def region_to_country(region: str) -> str:
    return REGION_TO_COUNTRY.get(str(region or '').strip(), '')


class _OriginSNIAdapter(HTTPAdapter):
    """直连 CZ1 源站 IP 同时使用真实域名 SNI 并严格校验 Let's Encrypt 证书。"""

    def __init__(self, server_hostname: str, **kwargs):
        self._server_hostname = server_hostname
        super().__init__(**kwargs)

    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        ctx = ssl.create_default_context()
        pool_kwargs['ssl_context'] = ctx
        pool_kwargs['server_hostname'] = self._server_hostname
        pool_kwargs['assert_hostname'] = self._server_hostname
        self.poolmanager = PoolManager(num_pools=connections, maxsize=maxsize, block=block, **pool_kwargs)


def _client():
    _load_env()
    token = os.environ.get('UMAMI_BRIDGE_TOKEN', '').strip()
    host = os.environ.get('UMAMI_BRIDGE_HOST', '').strip()
    sni = os.environ.get('UMAMI_BRIDGE_SNI', '').strip()
    if not token or not host or not sni:
        raise RuntimeError('未完整配置 UMAMI_BRIDGE_TOKEN / UMAMI_BRIDGE_HOST / UMAMI_BRIDGE_SNI')
    session = requests.Session()
    session.mount(f'https://{host}', _OriginSNIAdapter(sni))
    headers = {
        'Host': sni,
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
    }
    base_url = f'https://{host}'
    return session, base_url, headers


def _request(method: str, path: str, params=None, json_body=None, timeout: int = 20) -> dict:
    session, base_url, headers = _client()
    try:
        resp = session.request(
            method,
            base_url + path,
            params=params,
            json=json_body,
            headers=headers,
            timeout=timeout,
            verify=True,
        )
        data = resp.json()
        if resp.status_code >= 400 or not data.get('ok'):
            raise RuntimeError(data.get('error') or f'CZ1 桥接请求失败 HTTP {resp.status_code}')
        return data
    finally:
        session.close()


def lookup_site(url_or_domain: str) -> dict:
    dom = normalize_domain(url_or_domain)
    if not dom:
        return {'ok': False, 'found': False, 'error': '域名格式无效'}
    return _request('GET', '/api/overview/internal/site-lookup', params={'domain': dom}, timeout=12)


def provision_site(url_or_domain: str, manager: str = '', country: str = '',
                   umami_website_id: str = '', has_orders: bool = None,
                   order_count: int = None, in_order_system: bool = True) -> dict:
    dom = normalize_domain(url_or_domain)
    if not dom:
        return {'ok': False, 'error': '域名格式无效'}
    payload = {
        'domain': dom,
        'name': dom,
        'manager': (manager or '').strip(),
        'country': (country or '').strip().upper(),
        'umami_website_id': (umami_website_id or '').strip(),
        'in_order_system': bool(in_order_system),
    }
    if has_orders is not None:
        payload['has_orders'] = bool(has_orders)
    if order_count is not None:
        payload['order_count'] = int(order_count)
    return _request('POST', '/api/overview/internal/site-provision', json_body=payload, timeout=55)


def collect_hk_sites_payload(conn) -> tuple:
    """从 woo_analysis 收集所有活跃站点的负责人、国家、Umami ID 及真实出单数。"""
    sites_rows = conn.execute('''
        SELECT id, url, COALESCE(manager, '') AS manager, COALESCE(country, '') AS country,
               COALESCE(umami_website_id, '') AS umami_website_id
        FROM sites
        WHERE COALESCE(api_status, '') != 'archived'
        ORDER BY id
    ''').fetchall()
    order_rows = conn.execute('''
        SELECT source, COUNT(*) AS cnt
        FROM orders
        WHERE source IS NOT NULL AND source != ''
          AND status NOT IN ('checkout-draft', 'trash')
        GROUP BY source
    ''').fetchall()

    orders_by_dom = {}
    for row in order_rows:
        d = normalize_domain(row['source'])
        if d:
            orders_by_dom[d] = orders_by_dom.get(d, 0) + int(row['cnt'] or 0)

    payload_sites = []
    managers = set()
    for row in sites_rows:
        dom = normalize_domain(row['url'])
        if not dom:
            continue
        mgr = (row['manager'] or '').strip()
        if mgr:
            managers.add(mgr)
        cnt = orders_by_dom.get(dom, 0)
        payload_sites.append({
            'id': int(row['id']),
            'domain': dom,
            'url': row['url'],
            'manager': mgr,
            'country': (row['country'] or '').strip().upper(),
            'umami_website_id': (row['umami_website_id'] or '').strip(),
            'has_orders': cnt > 0,
            'order_count': cnt,
        })
    return payload_sites, sorted(managers)


def sync_all_sites_state(conn, force_meta: bool = False, only_domain: str = None) -> dict:
    """将香港站点状态、负责人及出单情况同步到 CZ1，并自动回填缺失的 umami_website_id。"""
    payload_sites, known_managers = collect_hk_sites_payload(conn)
    target_sites = payload_sites
    if only_domain:
        nd = normalize_domain(only_domain)
        target_sites = [s for s in payload_sites if s['domain'] == nd]
    res = _request(
        'POST',
        '/api/overview/internal/sync-state',
        json_body={
            'sites': target_sites,
            'known_managers': known_managers,
            'force_meta': bool(force_meta),
            'partial': bool(only_domain),
        },
        timeout=30,
    )
    matched = res.get('matched') or {}
    backfilled = 0
    for s in target_sites:
        info = matched.get(s['domain'])
        if not info:
            continue
        wid = (info.get('website_id') or '').strip()
        if wid and UUID_RE.fullmatch(wid) and wid != s['umami_website_id']:
            conn.execute('UPDATE sites SET umami_website_id = ? WHERE id = ?', (wid, s['id']))
            backfilled += 1
    if backfilled:
        conn.commit()
    return {
        'ok': True,
        'matched_count': len(matched),
        'backfilled_ids': backfilled,
        'total_umami': res.get('total_umami', 0),
        'matched': matched,
    }


def fetch_umami_daily_traffic(website_ids: list, start_date: str, end_date: str, timezone: str = 'Asia/Shanghai') -> dict:
    clean_ids = [str(w).strip().lower() for w in (website_ids or []) if UUID_RE.fullmatch(str(w or '').strip())]
    if not clean_ids:
        return {'ok': True, 'first_dates': {}, 'daily': {}}
    return _request(
        'POST',
        '/api/overview/internal/traffic-daily',
        json_body={
            'website_ids': sorted(set(clean_ids)),
            'start_date': str(start_date)[:10],
            'end_date': str(end_date)[:10],
            'timezone': timezone,
        },
        timeout=45,
    )
