#!/usr/bin/env python3
"""从历史 report_cache 抢救并补充固化因 51.la 4001 失效站点的日/月流量。"""
import calendar
import datetime as dt
import gzip
import json
import os
import pathlib
import re

import psycopg
from psycopg.rows import dict_row

APP_DIR = pathlib.Path('/www/wwwroot/woo-analysis')
ENV_PATH = pathlib.Path('/etc/woo-analysis/woo-analysis.env')
ARCHIVE_PATH = APP_DIR / 'backups' / '51la_frozen_history_20261120.json.gz'
RAW_TMP_PATH = APP_DIR / 'backups' / '51la_frozen_history_20261120.json.gz.raw_tmp'

LEGACY_DOMAIN_ALIAS = {
    'vapesklep.pl': 'vapeskleps.pl',
    'vapeskleps.pl': 'vapeskleps.pl',
    'vapego.pl': 'vapegopl.com',
    'vapegopl.com': 'vapegopl.com',
    'vapeszczyt.pl': 'vapeszczy.pl',
    'vapeszczy.pl': 'vapeszczy.pl',
    'vapeprime.pl': 'vapeprim.pl',
    'vapeprim.pl': 'vapeprim.pl',
    'ausmokeshop.com': 'ausmokeshop.com',
    'xvapehub.com': 'xvapehub.com',
}

KEY_RE = re.compile(r'(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})_(day|week|month)')
DAY_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')
MONTH_RE = re.compile(r'^\d{4}-\d{2}$')


def load_env(path: pathlib.Path):
    for line in path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if line and not line.startswith('#') and '=' in line:
            k, v = line.split('=', 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def is_full_month(ym: str, start_s: str, end_s: str) -> bool:
    if not (MONTH_RE.fullmatch(ym) and DAY_RE.fullmatch(start_s) and DAY_RE.fullmatch(end_s)):
        return False
    y, m = int(ym[:4]), int(ym[5:7])
    last_d = calendar.monthrange(y, m)[1]
    return start_s <= f'{ym}-01' and end_s >= f'{ym}-{last_d:02d}'


def main():
    load_env(ENV_PATH)
    if RAW_TMP_PATH.exists():
        RAW_TMP_PATH.unlink()

    with psycopg.connect(
        host=os.environ.get('WOO_DB_HOST', '127.0.0.1'),
        port=int(os.environ.get('WOO_DB_PORT', '5432')),
        dbname=os.environ.get('WOO_DB_NAME', 'woo_analysis'),
        user=os.environ.get('WOO_DB_USER'),
        password=os.environ.get('WOO_DB_PASSWORD'),
        autocommit=False,
    ) as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT cache_key, updated_at::text, cache_value FROM report_cache ORDER BY updated_at DESC;")
            cache_rows = cur.fetchall()

        day_candidates = {}    # (canon_dom, day_s) -> (uv, pv, new_u, br, key)
        month_candidates = {}  # (canon_dom, ym) -> (full_rank, uv, upd, pv, new_u, br, key)

        for cache_key, updated_at, cache_val in cache_rows:
            m = KEY_RE.search(cache_key or '')
            if not m:
                continue
            k_start, k_end, gran = m.groups()
            try:
                parsed = json.loads(cache_val)
            except Exception:
                continue
            dr = parsed.get('date_range') or {}
            if DAY_RE.fullmatch(str(dr.get('start') or '')):
                k_start = str(dr['start'])
            if DAY_RE.fullmatch(str(dr.get('end') or '')):
                k_end = str(dr['end'])
            data = parsed.get('data') or {}
            if not isinstance(data, dict):
                continue

            for raw_dom, canon_dom in LEGACY_DOMAIN_ALIAS.items():
                periods = data.get(raw_dom)
                if not isinstance(periods, dict):
                    continue
                for p_key, metrics in periods.items():
                    if not isinstance(metrics, dict):
                        continue
                    uv = metrics.get('uv')
                    pv = metrics.get('pv')
                    if uv is None and pv is None:
                        continue
                    uv_i = int(uv or 0)
                    pv_i = int(pv or 0)
                    if uv_i <= 0 and pv_i <= 0:
                        continue
                    new_u = int(metrics.get('newUserCount') or 0)
                    br = float(metrics.get('bounceRate') or 0.0)

                    if gran == 'day' and DAY_RE.fullmatch(str(p_key)):
                        dk = (canon_dom, p_key)
                        prev = day_candidates.get(dk)
                        if prev is None or (uv_i, pv_i) > (prev[0], prev[1]):
                            day_candidates[dk] = (uv_i, pv_i, new_u, br, cache_key)
                    elif gran == 'month' and MONTH_RE.fullmatch(str(p_key)):
                        full_rank = 1 if is_full_month(p_key, k_start, k_end) else 0
                        score = (full_rank, uv_i, str(updated_at or ''), pv_i)
                        mk = (canon_dom, p_key)
                        prev = month_candidates.get(mk)
                        if prev is None or score > prev[0]:
                            month_candidates[mk] = (score, uv_i, pv_i, new_u, br, cache_key)

        daily_inserts = [
            (dom, day_s, f'cache:{val[4][:64]}', val[0], val[1], val[2], val[3])
            for (dom, day_s), val in sorted(day_candidates.items())
        ]
        monthly_upserts = [
            (dom, ym, val[1], val[2], val[3], val[4], val[5])
            for (dom, ym), val in sorted(month_candidates.items())
        ]

        with conn.cursor() as cur:
            cur.executemany(
                """
                INSERT INTO traffic_daily_51la_frozen
                    (domain, stat_date, mask_id, uv, pv, new_user_count, bounce_rate)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (domain, stat_date) DO NOTHING;
                """,
                daily_inserts,
            )
            cur.executemany(
                """
                INSERT INTO traffic_monthly_51la_frozen
                    (domain, stat_month, uv, pv, new_user_count, bounce_rate, source_cache_key)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (domain, stat_month) DO UPDATE SET
                    uv = EXCLUDED.uv,
                    pv = EXCLUDED.pv,
                    new_user_count = EXCLUDED.new_user_count,
                    bounce_rate = EXCLUDED.bounce_rate,
                    source_cache_key = EXCLUDED.source_cache_key,
                    frozen_at = CURRENT_TIMESTAMP;
                """,
                monthly_upserts,
            )
        conn.commit()

        with conn.cursor(row_factory=dict_row) as cur:
            cur.execute("SELECT COUNT(DISTINCT domain) AS sites, COUNT(*) AS rows FROM traffic_daily_51la_frozen;")
            d_sum = dict(cur.fetchone())
            cur.execute("SELECT COUNT(DISTINCT domain) AS sites, COUNT(*) AS rows FROM traffic_monthly_51la_frozen;")
            m_sum = dict(cur.fetchone())
            cur.execute("""
                SELECT domain, COUNT(*) AS months, MIN(stat_month) AS min_m, MAX(stat_month) AS max_m, SUM(uv) AS sum_uv
                FROM traffic_monthly_51la_frozen
                GROUP BY domain ORDER BY sum_uv DESC;
            """)
            m_detail = [dict(r) for r in cur.fetchall()]
            cur.execute("""
                SELECT domain, stat_date::text AS stat_date, mask_id, uv, pv, new_user_count, sv, ip, bounce_rate, avg_duration
                FROM traffic_daily_51la_frozen ORDER BY domain, stat_date;
            """)
            all_daily = [dict(r) for r in cur.fetchall()]
            cur.execute("""
                SELECT domain, stat_month, uv, pv, new_user_count, bounce_rate, source_cache_key
                FROM traffic_monthly_51la_frozen ORDER BY domain, stat_month;
            """)
            all_monthly = [dict(r) for r in cur.fetchall()]

    with gzip.open(ARCHIVE_PATH, 'wt', encoding='utf-8') as gz:
        json.dump(
            {
                'frozen_at': dt.datetime.now().isoformat(),
                'daily_summary': d_sum,
                'monthly_summary': m_sum,
                'daily_rows': all_daily,
                'monthly_rows': all_monthly,
            },
            gz,
            ensure_ascii=False,
        )

    print('日固化总计:', d_sum, '| 新增候选:', len(daily_inserts))
    print('月补录总计:', m_sum)
    print('月补录各站详情:', m_detail)
    print('联合归档字节数:', ARCHIVE_PATH.stat().st_size)


if __name__ == '__main__':
    main()
