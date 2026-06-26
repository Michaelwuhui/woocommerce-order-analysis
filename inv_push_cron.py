"""
inv_push_cron.py — 独立的库存定时下推脚本(供 crontab 调用)。

与拉单核心(sync_utils.py / auto_sync.py)完全解耦:本脚本只读 inv_* 与 sites,
把本系统算出的可用库存下推到各站 WC,复用 Product Manager 的 PUT 白名单。

用法:
    venv/bin/python inv_push_cron.py            # 真实下推所有有映射的站点
    venv/bin/python inv_push_cron.py --dry      # 演练(只算不推)
    venv/bin/python inv_push_cron.py --site 2   # 仅某站点

建议 crontab(与 15 分钟拉单错峰,例如每 30 分钟):
    */30 * * * * cd /www/wwwroot/woo-analysis && venv/bin/python inv_push_cron.py >> inv_push.log 2>&1
"""

import sys
from inv_common import get_conn
import inv_push


def main():
    dry = '--dry' in sys.argv
    only_site = None
    if '--site' in sys.argv:
        try:
            only_site = int(sys.argv[sys.argv.index('--site') + 1])
        except (ValueError, IndexError):
            print('--site 需要一个站点 id'); return 1

    conn = get_conn()
    # 只推有映射的站点
    if only_site:
        site_ids = [only_site]
    else:
        site_ids = [r['site_id'] for r in conn.execute(
            'SELECT DISTINCT site_id FROM inv_site_sku_map WHERE is_active=1').fetchall()]
    if not site_ids:
        print('没有任何站点配置了 SKU 映射,跳过。')
        conn.close(); return 0

    total_ok = total_err = total = 0
    for sid in site_ids:
        res = inv_push.push_site(conn, sid, dry_run=dry)
        total += res['total']; total_ok += res['ok']; total_err += res['error']
        tag = '[演练]' if dry else ''
        print(f"{tag} 站点 {sid} {res.get('source','')}: 商品 {res['total']} 推送ok {res['ok']} 失败 {res['error']}"
              + (f" 致命:{res['fatal']}" if res.get('fatal') else ''))
    conn.close()
    print(f"完成{'(演练)' if dry else ''}:站点 {len(site_ids)} 个,商品 {total},成功 {total_ok},失败 {total_err}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
