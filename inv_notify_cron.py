"""
inv_notify_cron.py — 独立的库存通知扫描脚本(供 crontab 调用)。

扫描补货/临期过期/滞销,写站内通知;若配置了 SMTP(settings 表 inv_smtp_*)则发邮件。
与拉单核心解耦。建议每天跑 1-2 次:
    0 9 * * * cd /www/wwwroot/woo-analysis && venv/bin/python inv_notify_cron.py >> inv_notify.log 2>&1
"""

import sys
from inv_common import get_conn
import inv_notify


def main():
    conn = get_conn()
    counts = inv_notify.generate_all(conn)
    emailed = inv_notify.send_pending_emails(conn)
    conn.close()
    print(f"通知扫描完成:补货 {counts['restock']} / 临期过期 {counts['expiry']} / 滞销 {counts['slow_moving']}"
          f" · 邮件 {('未配置/失败' if emailed < 0 or emailed == 0 else str(emailed)+'封')}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
