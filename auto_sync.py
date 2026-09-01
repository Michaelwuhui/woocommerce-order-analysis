#!/usr/bin/env python3
"""
Auto-sync script for cron execution.
This script is called by cron to automatically sync all sites.
Runs sites sequentially with incremental checkpoints and bounded note refreshes.
"""
import sqlite3
import threading
import fcntl
import time
from datetime import datetime
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import sync_utils

DB_FILE = 'woocommerce_orders.db'
MAX_SITE_CONCURRENCY = 1
INTER_SITE_DELAY_SECONDS = 1
INCREMENTAL_OVERLAP_MINUTES = 10
NOTES_ACTIVE_LIMIT = 25
NOTES_REFRESH_INTERVAL_HOURS = 24
NOTE_WORKERS = 1
LOCK_FILE = '/tmp/woo-analysis-auto-sync.lock'

# 线程安全的打印锁
_print_lock = threading.Lock()

def safe_print(msg):
    """线程安全的打印"""
    with _print_lock:
        print(msg, flush=True)

def get_db_connection():
    """Create database connection"""
    conn = sqlite3.connect(DB_FILE, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=30000")
    return conn

def log_sync(site_id, site_url, status, message, new_orders=0, updated_orders=0, duration=0):
    """Log sync result to database"""
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO sync_logs (site_id, site_url, status, message, new_orders, updated_orders, sync_time, duration_seconds)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (site_id, site_url, status, message, new_orders, updated_orders, 
          datetime.now().strftime('%Y-%m-%d %H:%M:%S'), duration))
    conn.commit()
    conn.close()

def sync_one_site(site):
    """同步单个站点（在线程中执行）"""
    site_id = site['id']
    site_url = site['url']
    consumer_key = site['consumer_key']
    consumer_secret = site['consumer_secret']
    
    safe_print(f"[Site] Syncing: {site_url}")
    
    start_time = datetime.now()
    
    def progress_callback(msg):
        safe_print(f"  [{site_url}] {msg}")
    
    try:
        result = sync_utils.sync_site(
            site_url,
            consumer_key,
            consumer_secret,
            progress_callback,
            sync_days=0,
            incremental_overlap_minutes=INCREMENTAL_OVERLAP_MINUTES,
            notes_active_limit=NOTES_ACTIVE_LIMIT,
            notes_refresh_interval_hours=NOTES_REFRESH_INTERVAL_HOURS,
            note_workers=NOTE_WORKERS,
        )
        
        duration = int((datetime.now() - start_time).total_seconds())
        
        if result['status'] == 'success':
            # Update last_sync time
            conn = get_db_connection()
            conn.execute('UPDATE sites SET last_sync = ? WHERE id = ?',
                        (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), site_id))
            conn.commit()
            conn.close()
            
            log_sync(
                site_id, site_url, 'success',
                f"Synced successfully: {result.get('new_orders', 0)} new, "
                f"{result.get('updated_orders', 0)} updated, "
                f"{result.get('notes_checked', 0)} notes checked, "
                f"{result.get('note_failures', 0)} note failures",
                result.get('new_orders', 0), result.get('updated_orders', 0), duration
            )
            safe_print(f"  [{site_url}] ✓ Success ({duration}s)")
        else:
            log_sync(site_id, site_url, 'error', result.get('message', 'Unknown error'), 0, 0, duration)
            safe_print(f"  [{site_url}] ✗ Error: {result.get('message', 'Unknown error')}")
            
        return {'site_url': site_url, 'status': result['status'], 'duration': duration}
        
    except Exception as e:
        duration = int((datetime.now() - start_time).total_seconds())
        log_sync(site_id, site_url, 'error', str(e), 0, 0, duration)
        safe_print(f"  [{site_url}] ✗ Exception: {e}")
        return {'site_url': site_url, 'status': 'error', 'duration': duration}

def run_sync_batch():
    safe_print(
        f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
        f"Auto-sync started (site concurrency {MAX_SITE_CONCURRENCY})"
    )
    
    conn = get_db_connection()
    sites = conn.execute('SELECT * FROM sites').fetchall()
    conn.close()
    
    if not sites:
        safe_print("No sites configured")
        return
    
    safe_print(f"Total sites: {len(sites)}")
    
    results = []
    for index, site in enumerate(sites):
        try:
            results.append(sync_one_site(site))
        except Exception as e:
            safe_print(f"  [{site['url']}] ✗ Sync exception: {e}")
            results.append({'site_url': site['url'], 'status': 'error', 'duration': 0})
        if index + 1 < len(sites):
            time.sleep(INTER_SITE_DELAY_SECONDS)
    
    # 输出汇总
    success_count = sum(1 for r in results if r['status'] == 'success')
    error_count = sum(1 for r in results if r['status'] != 'success')
    total_max_duration = max(r['duration'] for r in results) if results else 0
    
    safe_print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Auto-sync completed: "
               f"{success_count} success, {error_count} errors, "
               f"longest site took {total_max_duration}s")

    # Enforce the customer blocklist (auto-cancel blacklisted COD orders).
    # Isolated in its own try so a failure here never affects sync.
    try:
        import blocklist
        conn = get_db_connection()
        try:
            if blocklist.is_globally_enabled(conn):
                bl = blocklist.enforce(conn, progress=safe_print, actor='auto-sync')
                safe_print(f"[blocklist] checked={bl['checked']} cancelled={bl['cancelled']} "
                           f"errors={bl['errors']} skipped={bl['skipped']}"
                           + ("  [ABORTED-safety-cap]" if bl.get('aborted') else ""))
            else:
                safe_print("[blocklist] globally disabled — skipped")
        finally:
            conn.close()
    except Exception as e:
        safe_print(f"[blocklist] enforcement failed: {e}")

    # Auto-confirm carrier-delivered COD orders (待确认结局 → 已签收), if enabled.
    # Isolated in its own try so a failure here never affects sync.
    try:
        import auto_confirm
        conn = get_db_connection()
        try:
            if auto_confirm.is_enabled(conn):
                ac = auto_confirm.enforce(conn, progress=safe_print, actor='auto-sync')
                safe_print(f"[auto-confirm] checked={ac['checked']} confirmed={ac['confirmed']} "
                           f"(synced={ac['synced']} local_only={ac['local_only']}) "
                           f"errors={ac['errors']} deferred={ac['capped']}")
            else:
                safe_print("[auto-confirm] disabled — skipped")
        finally:
            conn.close()
    except Exception as e:
        safe_print(f"[auto-confirm] enforcement failed: {e}")


def main():
    """Run one batch only; skip safely if an earlier batch is still active."""
    lock_handle = open(LOCK_FILE, 'w')
    try:
        fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        safe_print(
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
            "Auto-sync skipped: another batch is still running"
        )
        lock_handle.close()
        return

    try:
        run_sync_batch()
    finally:
        fcntl.flock(lock_handle.fileno(), fcntl.LOCK_UN)
        lock_handle.close()


if __name__ == '__main__':
    main()
