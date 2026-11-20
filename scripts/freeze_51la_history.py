#!/usr/bin/env python3
"""将 51.la 历史逐日流量永久固化到 PostgreSQL 表 traffic_daily_51la_frozen。"""
import ast
import concurrent.futures
import datetime as dt
import gzip
import json
import os
import pathlib
import random
import re
import string
import sys
import threading
import time

import psycopg
from psycopg.rows import dict_row
import requests

APP_DIR = pathlib.Path('/www/wwwroot/woo-analysis')
ENV_PATH = pathlib.Path('/etc/woo-analysis/woo-analysis.env')
ARCHIVE_PATH = APP_DIR / 'backups' / '51la_frozen_history_20261120.json.gz'
_tls = threading.local()


def load_env(path: pathlib.Path):
    if not path.is_file():
        return
    for line in path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        k, v = line.split('=', 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def normalize_domain(url: str) -> str:
    if not url:
        return ''
    u = url.lower().strip()
    u = re.sub(r'^https?://', '', u)
    u = u.split('/')[0].split(':')[0]
    if u.startswith('www.'):
        u = u[4:]
    return u


def load_la_config(app_py: pathlib.Path) -> dict:
    src = app_py.read_text(encoding='utf-8')
    tree = ast.parse(src)
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == 'LA_API_CONFIG':
                    return ast.literal_eval(node.value)
    raise RuntimeError('未能在 app.py 中找到 LA_API_CONFIG')


def iter_windows(start_day: dt.date, end_day: dt.date, span_days: int = 90):
    cur = start_day
    while cur <= end_day:
        nxt = min(end_day, cur + dt.timedelta(days=span_days - 1))
        yield cur.isoformat(), nxt.isoformat()
        cur = nxt + dt.timedelta(days=1)


def _session() -> requests.Session:
    s = getattr(_tls, 'session', None)
    if s is None:
        s = requests.Session()
        s.headers.update({'Content-Type': 'application/json'})
        _tls.session = s
    return s


def fetch_chunk(api_url: str, access_key: str, mask_id: str, start_s: str, end_s: str):
    session = _session()
    for attempt in range(3):
        nonce = ''.join(random.choices(string.ascii_letters + string.digits, k=4))
        ts = str(int(time.time() * 1000))
        payload = {
            'maskId': mask_id,
            'startDay': start_s,
            'endDay': end_s,
            'accessKey': access_key,
            'nonce': nonce,
            'timestamp': ts,
            'sign': access_key,
        }
        try:
            resp = session.post(api_url, json=payload, timeout=15)
            data = resp.json()
            if data.get('success') and data.get('code') == '0000':
                return data.get('data') or []
            code_s = str(data.get('code') or '')
            msg = str(data.get('message') or code_s or 'unknown')
            if code_s == '4001':
                return {'_skipped_4001': True}
            if attempt == 2:
                raise RuntimeError(f'51.la 返回错误: {msg}')
        except Exception as exc:
            if attempt == 2:
                raise RuntimeError(f'51.la 请求异常: {exc}') from exc
        time.sleep(0.6 * (attempt + 1))
    return []


def fetch_site_all(dom: str, mid: str, windows: list, api_url: str, access_key: str):
    site_days = {}
    errs = []
    for ws, we in windows:
        try:
            items = fetch_chunk(api_url, access_key, mid, ws, we)
            if isinstance(items, dict) and items.get('_skipped_4001'):
                errs.append(('SKIP_4001', dom, ws, we))
                break
            for it in items:
                t = str(it.get('time') or '')[:10]
                if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', t):
                    continue
                site_days[t] = (
                    dom, t, mid,
                    int(it.get('uv') or 0),
                    int(it.get('pv') or 0),
                    int(it.get('newUserCount') or 0),
                    int(it.get('sv') or 0),
                    int(it.get('ip') or 0),
                    float(it.get('bounceRate') or 0.0),
                    float(it.get('avgDuration') or 0.0),
                )
        except Exception as exc:
            errs.append((dom, ws, we, str(exc)))
        time.sleep(0.05)
    return dom, list(site_days.values()), errs


def main():
    load_env(ENV_PATH)
    cfg = load_la_config(APP_DIR / 'app.py')
    api_url = cfg['api_url']
    access_key = cfg['access_key']

    conn_kw = dict(
        host=os.environ.get('WOO_DB_HOST', '127.0.0.1'),
        port=int(os.environ.get('WOO_DB_PORT', '5432')),
        dbname=os.environ.get('WOO_DB_NAME', 'woo_analysis'),
        user=os.environ.get('WOO_DB_USER', 'postgres'),
        password=os.environ.get('WOO_DB_PASSWORD', ''),
        autocommit=False,
    )

    with psycopg.connect(**conn_kw) as conn:
        domain_masks = {}
        with conn.cursor() as cur:
            cur.execute("SELECT url, mask_id FROM sites WHERE mask_id IS NOT NULL AND TRIM(mask_id) != '' ORDER BY id;")
            for url, mid in cur.fetchall():
                dom = normalize_domain(url)
                if dom:
                    domain_masks[dom] = mid.strip()
    for legacy_dom, mid in (cfg.get('sites') or {}).items():
        dom = normalize_domain(legacy_dom)
        if dom and dom not in domain_masks:
            domain_masks[dom] = str(mid).strip()

    start_day = dt.date(2025, 6, 1)
    end_day = dt.date(2026, 11, 20)
    windows = list(iter_windows(start_day, end_day, 90))
    items = sorted(domain_masks.items())
    print(f'待固化站点数量: {len(items)}, 时间窗数量: {len(windows)} ({start_day} ~ {end_day})', flush=True)

    rows_to_upsert = []
    errors = []
    done_cnt = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
        futs = [pool.submit(fetch_site_all, dom, mid, windows, api_url, access_key) for dom, mid in items]
        for fut in concurrent.futures.as_completed(futs):
            dom, rows, errs = fut.result()
            rows_to_upsert.extend(rows)
            errors.extend(errs)
            done_cnt += 1
            if done_cnt % 10 == 0 or done_cnt == len(items):
                print(f'抓取进度: {done_cnt}/{len(items)} 站 (累计 {len(rows_to_upsert)} 日行)', flush=True)

    tmp_gz = ARCHIVE_PATH.with_name(ARCHIVE_PATH.name + '.raw_tmp')
    with gzip.open(tmp_gz, 'wt', encoding='utf-8') as gz:
        json.dump({'raw_count': len(rows_to_upsert), 'rows': [list(r) for r in rows_to_upsert]}, gz, ensure_ascii=False)
    print('原始抓取已暂存:', tmp_gz, 'bytes:', tmp_gz.stat().st_size, flush=True)

    with psycopg.connect(**conn_kw) as conn:
        with conn.cursor() as cur:
            cur.executemany(
                """
                INSERT INTO traffic_daily_51la_frozen
                    (domain, stat_date, mask_id, uv, pv, new_user_count, sv, ip, bounce_rate, avg_duration)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (domain, stat_date) DO UPDATE SET
                    mask_id = EXCLUDED.mask_id,
                    uv = EXCLUDED.uv,
                    pv = EXCLUDED.pv,
                    new_user_count = EXCLUDED.new_user_count,
                    sv = EXCLUDED.sv,
                    ip = EXCLUDED.ip,
                    bounce_rate = EXCLUDED.bounce_rate,
                    avg_duration = EXCLUDED.avg_duration,
                    frozen_at = CURRENT_TIMESTAMP;
                """,
                rows_to_upsert,
            )
        conn.commit()

        with conn.cursor(row_factory=dict_row) as cur:
            cur.execute("""
                SELECT domain, stat_date::text AS stat_date, mask_id, uv, pv, new_user_count, sv, ip, bounce_rate, avg_duration
                FROM traffic_daily_51la_frozen
                ORDER BY domain, stat_date;
            """)
            all_rows = [dict(r) for r in cur.fetchall()]
            cur.execute("""
                SELECT COUNT(DISTINCT domain) AS sites, COUNT(*) AS rows, MIN(stat_date)::text AS min_d, MAX(stat_date)::text AS max_d
                FROM traffic_daily_51la_frozen;
            """)
            summary = dict(cur.fetchone())
            cur.execute("""
                SELECT domain, SUM(uv) AS total_uv, SUM(pv) AS total_pv, MIN(stat_date)::text AS s, MAX(stat_date)::text AS e
                FROM traffic_daily_51la_frozen
                WHERE domain IN ('strefajednorazowek.com', 'vapicoaus.com', 'vapeskleps.pl', 'strefajednorazowek.pl', 'buchmistrz.pl')
                GROUP BY domain ORDER BY domain;
            """)
            samples = [dict(r) for r in cur.fetchall()]

    with gzip.open(ARCHIVE_PATH, 'wt', encoding='utf-8') as gz:
        json.dump({'frozen_at': dt.datetime.now().isoformat(), 'summary': summary, 'rows': all_rows}, gz, ensure_ascii=False)

    print('固化汇总:', summary)
    print('核心站抽样:', samples)
    print('归档大小(bytes):', ARCHIVE_PATH.stat().st_size)
    skipped = [e for e in errors if e and e[0] == 'SKIP_4001']
    real_errs = [e for e in errors if not (e and e[0] == 'SKIP_4001')]
    if skipped:
        print(f'SKIP_4001_SITES={len(skipped)}:', sorted({e[1] for e in skipped}))
    if real_errs:
        print(f'ERROR_COUNT={len(real_errs)}:', real_errs[:10], file=sys.stderr)
        sys.exit(2)


if __name__ == '__main__':
    main()
