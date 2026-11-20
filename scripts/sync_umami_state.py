#!/usr/bin/env python3
"""定期将香港订单系统的出单状态、负责人与国家同步到 CZ1 Umami 并回填 Website ID。"""
import json
import os
import pathlib
import sys

APP_DIR = pathlib.Path('/www/wwwroot/woo-analysis')
sys.path.insert(0, str(APP_DIR))


def load_env(path: pathlib.Path):
    if not path.is_file():
        return
    for line in path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if line and not line.startswith('#') and '=' in line:
            k, v = line.split('=', 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def main():
    load_env(pathlib.Path('/etc/woo-analysis/woo-analysis.env'))
    load_env(pathlib.Path('/etc/woo-analysis/umami-sync.env'))
    import db_backend
    import umami_bridge

    conn = db_backend.connect()
    try:
        res = umami_bridge.sync_all_sites_state(conn, force_meta=False)
        print(json.dumps({
            'ok': res.get('ok'),
            'matched_count': res.get('matched_count'),
            'backfilled_ids': res.get('backfilled_ids'),
            'total_umami': res.get('total_umami'),
        }, ensure_ascii=False))
    finally:
        conn.close()


if __name__ == '__main__':
    main()
