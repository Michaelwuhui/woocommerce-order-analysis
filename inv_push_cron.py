"""Run due per-site Woo inventory sync jobs.

Safe defaults:
  * the global switch must be enabled;
  * each site must be observe or live;
  * observe never writes WooCommerce;
  * --site is a dry-run unless --live is also supplied.

Suggested cron (the application-level site lock prevents overlap):
  */5 * * * * cd /www/wwwroot/woo-analysis && flock -n /tmp/woo-inv-push.lock \
    venv/bin/python inv_push_cron.py >> inv_push_cron.log 2>&1
"""

import argparse
import sys

from inv_common import get_conn
import inv_push


def _args(argv=None):
    parser = argparse.ArgumentParser(description="Woo 库存按站点自动同步")
    parser.add_argument("--site", type=int, help="手动运行一个站点")
    parser.add_argument(
        "--live",
        action="store_true",
        help="与 --site 同用时允许真实写入；缺省只演练",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = _args(argv)
    conn = get_conn()
    try:
        if args.site:
            site_ids = [args.site]
            trigger = "manual_cli"
        else:
            site_ids = inv_push.scheduler_site_ids(conn)
            trigger = "scheduler"
        if not site_ids:
            print("没有到期且已启用的站点，跳过。")
            return 0

        failed = 0
        for site_id in site_ids:
            force_dry = (not args.live) if args.site else None
            try:
                result = inv_push.execute_site_sync(
                    conn,
                    site_id,
                    trigger_type=trigger,
                    force_dry_run=force_dry,
                    operator=(None, "system:auto_inventory_push"),
                )
            except Exception as exc:
                failed += 1
                print(f"站点 {site_id}: 未处理异常: {exc}")
                continue
            if result.get("status") == "skipped":
                print(f"站点 {site_id}: 跳过 - {result.get('reason', '')}")
                continue
            failed += 1 if result.get("status") in ("error", "partial") else 0
            tag = "演练" if result.get("dry_run") else "正式"
            print(
                f"站点 {site_id} [{tag}] {result.get('status')}: "
                f"商品 {result.get('total', 0)}，写入 {result.get('ok', 0)}，"
                f"未变化 {result.get('unchanged', 0)}，失败 {result.get('error', 0)}"
                + (f"，致命: {result['fatal']}" if result.get("fatal") else "")
            )
        return 1 if failed else 0
    finally:
        conn.close()


if __name__ == "__main__":
    sys.exit(main())
