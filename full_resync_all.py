#!/usr/bin/env python3
"""Fetch every historical order page for every configured site."""

import argparse
import os
import random
import sqlite3
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import sync_utils
from sync_process_lock import SyncAlreadyRunning, exclusive_sync_lock
from sync_runtime_status import init_sync_runtime_status, save_sync_runtime_status
from woocommerce import API


DB_FILE = "woocommerce_orders.db"
MAX_PAGE_ATTEMPTS = 3
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}
MAX_STATUS_LOGS = 200


def _connect_db():
    conn = sqlite3.connect(DB_FILE, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


class RuntimeStatusReporter:
    """Persist progress so every Gunicorn worker can serve the same status."""

    def __init__(self, status_id):
        self.status_id = int(status_id) if status_id is not None else None
        self.entry = {
            "status": "running",
            "message": "正在启动全量深度同步...",
            "logs": [],
            "progress": 0,
        }
        if self.status_id is not None:
            conn = _connect_db()
            try:
                init_sync_runtime_status(conn)
            finally:
                conn.close()

    def update(self, *, status=None, message=None, progress=None, log=None):
        if status is not None:
            self.entry["status"] = status
        if message is not None:
            self.entry["message"] = message
        if progress is not None:
            self.entry["progress"] = max(0.0, min(100.0, float(progress)))
        if log:
            self.entry["logs"].append(str(log))
            self.entry["logs"] = self.entry["logs"][-MAX_STATUS_LOGS:]
        if self.status_id is None:
            return
        conn = _connect_db()
        try:
            save_sync_runtime_status(conn, self.status_id, self.entry)
        finally:
            conn.close()


def fetch_orders_page(
    wcapi,
    page,
    *,
    max_attempts=MAX_PAGE_ATTEMPTS,
    sleep_fn=time.sleep,
    log_fn=print,
):
    """Fetch one order page with bounded retries and exponential backoff."""
    last_exception = None
    for attempt in range(1, max_attempts + 1):
        try:
            response = wcapi.get(
                "orders",
                params={
                    "per_page": 100,
                    "page": page,
                    "expand": (
                        "line_items,shipping_lines,tax_lines,fee_lines,"
                        "coupon_lines,refunds"
                    ),
                },
            )
        except Exception as exc:
            last_exception = exc
            if attempt >= max_attempts:
                raise
            delay = 2 ** (attempt - 1)
            log_fn(
                f"  Page {page}: request failed ({attempt}/{max_attempts}): "
                f"{exc}; retrying in {delay}s"
            )
            sleep_fn(delay)
            continue

        if response.status_code not in RETRYABLE_HTTP_STATUS:
            return response
        if attempt >= max_attempts:
            return response
        delay = 2 ** (attempt - 1)
        log_fn(
            f"  Page {page}: HTTP {response.status_code} "
            f"({attempt}/{max_attempts}); retrying in {delay}s"
        )
        sleep_fn(delay)

    if last_exception is not None:
        raise last_exception
    raise RuntimeError(f"page {page} request did not return a response")


def _count(url):
    conn = _connect_db()
    try:
        return conn.execute(
            "SELECT COUNT(*) FROM orders WHERE source = ?", (url,)
        ).fetchone()[0]
    finally:
        conn.close()


def run_full_resync(reporter):
    conn = _connect_db()
    try:
        sites = conn.execute(
            "SELECT id, url, consumer_key, consumer_secret FROM sites "
            "WHERE consumer_key IS NOT NULL AND consumer_key != ''"
        ).fetchall()
    finally:
        conn.close()

    site_count = len(sites)
    reporter.update(
        message=f"准备同步 {site_count} 个站点",
        progress=0,
        log=f"Found {site_count} configured sites",
    )
    summary = []
    failed_sites = 0

    for site_index, site in enumerate(sites, start=1):
        url = site["url"]
        print("\n=========================", flush=True)
        print(f">>> Full resync: {url}", flush=True)
        print("=========================", flush=True)
        before = _count(url)
        total = 0
        page = 1
        site_failed = False
        reporter.update(
            message=f"正在同步站点 {site_index}/{site_count}: {url}",
            progress=((site_index - 1) / site_count * 100) if site_count else 0,
            log=f"[{site_index}/{site_count}] Started {url}",
        )

        try:
            wcapi = API(
                url=url,
                consumer_key=site["consumer_key"],
                consumer_secret=site["consumer_secret"],
                version="wc/v3",
                timeout=30,
            )
            while True:
                time.sleep(random.uniform(0.2, 0.4))
                reporter.update(
                    message=f"站点 {site_index}/{site_count}，正在请求第 {page} 页",
                )

                def report_retry(line):
                    print(line, flush=True)
                    reporter.update(message=line.strip(), log=line.strip())

                try:
                    response = fetch_orders_page(
                        wcapi,
                        page,
                        log_fn=report_retry,
                    )
                except Exception as exc:
                    message = (
                        f"{url} 第 {page} 页请求在 {MAX_PAGE_ATTEMPTS} 次后失败: {exc}"
                    )
                    print(f">>> ERROR: {message}", flush=True)
                    reporter.update(message=message, log=message)
                    site_failed = True
                    break

                if response.status_code != 200:
                    message = (
                        f"{url} 第 {page} 页 HTTP {response.status_code}: "
                        f"{response.text[:200]}"
                    )
                    print(f">>> ERROR: {message}", flush=True)
                    reporter.update(message=message, log=message)
                    site_failed = True
                    break

                data = response.json()
                if not data:
                    print(f"  Page {page}: empty, finished", flush=True)
                    break
                for order in data:
                    order["source"] = url
                sync_utils.save_orders_to_db(data)
                total += len(data)
                total_pages = int(response.headers.get("X-WP-TotalPages") or 0)
                page_fraction = min(1.0, page / total_pages) if total_pages else 0.5
                progress = (
                    ((site_index - 1) + page_fraction) / site_count * 100
                    if site_count
                    else 0
                )
                first_date = data[0].get("date_created", "")
                last_date = data[-1].get("date_created", "")
                line = (
                    f"[{site_index}/{site_count}] {url} page {page}"
                    f"/{total_pages or '?'}: {len(data)} orders, total {total}"
                )
                print(
                    f"  Page {page}: {len(data)} orders "
                    f"({first_date} ~ {last_date}), running total {total}",
                    flush=True,
                )
                reporter.update(
                    message=(
                        f"站点 {site_index}/{site_count}，第 {page}"
                        f"/{total_pages or '?'} 页，已读取 {total} 单"
                    ),
                    progress=progress,
                    log=line,
                )
                if total_pages and page >= total_pages:
                    print(f"  Reached final page {total_pages}", flush=True)
                    break
                page += 1

        except Exception as exc:
            message = f"{url} 同步失败: {exc}"
            print(f">>> ERROR: {message}", flush=True)
            reporter.update(message=message, log=message)
            site_failed = True

        after = _count(url)
        summary.append((url, before, after, total, site_failed))
        if site_failed:
            failed_sites += 1
        reporter.update(
            message=f"已处理站点 {site_index}/{site_count}: {url}",
            progress=(site_index / site_count * 100) if site_count else 100,
            log=(
                f"[{site_index}/{site_count}] Finished {url}: "
                f"DB {before}->{after}, fetched {total}"
            ),
        )
        print(
            f">>> {url}: DB before={before}, after={after}, fetched={total}",
            flush=True,
        )

    print("\n\n=== Summary ===", flush=True)
    for url, before, after, fetched, site_failed in summary:
        delta = after - before
        flag = "ERROR" if site_failed else "OK"
        print(
            f"  {flag} {delta:+7}  {url}  "
            f"({before} -> {after}, fetched {fetched})",
            flush=True,
        )

    if failed_sites:
        reporter.update(
            status="error",
            message=f"深度同步完成，但 {failed_sites}/{site_count} 个站点失败",
            progress=100,
            log=f"Completed with {failed_sites} failed sites",
        )
        return 1
    reporter.update(
        status="success",
        message=f"深度同步完成：{site_count} 个站点",
        progress=100,
        log="Completed successfully",
    )
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--status-id", type=int)
    args = parser.parse_args(argv)
    reporter = RuntimeStatusReporter(args.status_id)
    try:
        with exclusive_sync_lock():
            reporter.update(log="Exclusive synchronization lock acquired")
            return run_full_resync(reporter)
    except SyncAlreadyRunning:
        message = "已有自动同步或全量同步正在运行，本次深度同步未启动"
        print(message, flush=True)
        reporter.update(status="error", message=message, progress=0, log=message)
        return 75


if __name__ == "__main__":
    raise SystemExit(main())
