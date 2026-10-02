"""Persist temporary store outages without changing site or order identity."""
from datetime import datetime, timezone
from zoneinfo import ZoneInfo


FAILURE_LABELS = {
    "dns": "域名 DNS 解析失败",
    "tls": "站点 HTTPS / 源站证书连接失败",
    "timeout": "站点响应超时",
    "connection": "无法连接站点",
    "http": "站点服务暂不可用",
    "redirect": "站点循环跳转，请检查原域名的站点路由",
    "html": "订单接口返回网页，可能停站或被访问规则拦截",
    "json": "订单接口返回无效 JSON",
}


def backoff_seconds(failure_count: int) -> int:
    return min(3600, 300 * (2 ** min(4, max(0, int(failure_count) - 1))))


def availability_message(kind: str, next_check_at) -> str:
    label = FAILURE_LABELS.get(kind, "站点暂不可用")
    stamp = next_check_at
    if not isinstance(stamp, datetime):
        stamp = datetime.fromisoformat(str(stamp).replace("Z", "+00:00"))
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=timezone.utc)
    local = stamp.astimezone(ZoneInfo("Asia/Hong_Kong"))
    return (
        f"{label}；已自动暂缓，最早 {local:%m-%d %H:%M}（北京时间）复查，"
        "恢复后自动补同步。也可点击快速同步提前检查。"
    )


def load_site_health(connection, site_ids) -> dict[int, dict]:
    ids = list(dict.fromkeys(int(value) for value in site_ids))
    if not ids:
        return {}
    placeholders = ",".join("?" for _ in ids)
    rows = connection.execute(
        f"""SELECT site_id,failure_kind,failure_count,next_check_at,last_error,
                   last_failure_at,next_check_at>CURRENT_TIMESTAMP AS cooling_down
            FROM sync_site_health WHERE site_id IN ({placeholders})""",
        tuple(ids),
    ).fetchall()
    result = {}
    for row in rows:
        item = dict(row)
        item["availability_message"] = availability_message(
            item["failure_kind"], item["next_check_at"]
        )
        for key in ("next_check_at", "last_failure_at"):
            item[key] = item[key].isoformat() if isinstance(item[key], datetime) else str(item[key])
        result[int(item["site_id"])] = item
    return result


def record_site_failure(connection, site_id: int, run_id: str, kind: str, message: str) -> dict:
    if kind not in FAILURE_LABELS:
        raise ValueError("unknown temporary site failure")
    previous = connection.execute(
        "SELECT failure_count,last_failure_run_id FROM sync_site_health WHERE site_id=? FOR UPDATE",
        (site_id,),
    ).fetchone()
    if previous and str(previous["last_failure_run_id"]) == str(run_id):
        return load_site_health(connection, [site_id])[site_id]
    count = int(previous["failure_count"]) + 1 if previous else 1
    connection.execute(
        """INSERT INTO sync_site_health
                (site_id,failure_kind,failure_count,next_check_at,last_error,
                 last_failure_at,last_failure_run_id)
            VALUES (?,?,?,CURRENT_TIMESTAMP + (? * interval '1 second'),?,CURRENT_TIMESTAMP,?)
            ON CONFLICT(site_id) DO UPDATE SET
                failure_kind=excluded.failure_kind,failure_count=excluded.failure_count,
                next_check_at=excluded.next_check_at,last_error=excluded.last_error,
                last_failure_at=excluded.last_failure_at,last_failure_run_id=excluded.last_failure_run_id""",
        (site_id, kind, count, backoff_seconds(count), str(message)[:2000], run_id),
    )
    return load_site_health(connection, [site_id])[site_id]


def clear_site_failure(connection, site_id: int) -> bool:
    return connection.execute(
        "DELETE FROM sync_site_health WHERE site_id=? RETURNING site_id", (site_id,)
    ).fetchone() is not None


def due_site_rechecks(connection) -> list[int]:
    return [int(row[0]) for row in connection.execute(
        "SELECT site_id FROM sync_site_health WHERE next_check_at<=CURRENT_TIMESTAMP "
        "ORDER BY next_check_at,site_id LIMIT 100"
    ).fetchall()]
