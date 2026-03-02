#!/usr/bin/env python3
"""
Auto-sync script for cron execution.
This script is called by cron to automatically sync all sites.
Optimized: uses ThreadPoolExecutor for concurrent sync (max 4 workers).
"""
import sqlite3
import json
import threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import sync_utils

DB_FILE = 'woocommerce_orders.db'
MAX_WORKERS = 4

# 线程安全的打印锁
_print_lock = threading.Lock()

def safe_print(msg):
    """线程安全的打印"""
    with _print_lock:
        print(msg, flush=True)

def get_db_connection():
    """Create database connection"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
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
    
    safe_print(f"[Thread] Syncing: {site_url}")
    
    start_time = datetime.now()
    
    def progress_callback(msg):
        safe_print(f"  [{site_url}] {msg}")
    
    try:
        result = sync_utils.sync_site(
            site_url,
            consumer_key,
            consumer_secret,
            progress_callback
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
                f"Synced successfully: {result.get('new_orders', 0)} new, {result.get('updated_orders', 0)} updated",
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

def main():
    safe_print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Auto-sync started (max {MAX_WORKERS} concurrent)")
    
    conn = get_db_connection()
    sites = conn.execute('SELECT * FROM sites').fetchall()
    conn.close()
    
    if not sites:
        safe_print("No sites configured")
        return
    
    safe_print(f"Total sites: {len(sites)}")
    
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(sync_one_site, site): site for site in sites}
        for future in as_completed(futures):
            try:
                result = future.result()
                results.append(result)
            except Exception as e:
                site = futures[future]
                safe_print(f"  [{site['url']}] ✗ Future exception: {e}")
                results.append({'site_url': site['url'], 'status': 'error', 'duration': 0})
    
    # 输出汇总
    success_count = sum(1 for r in results if r['status'] == 'success')
    error_count = sum(1 for r in results if r['status'] != 'success')
    total_max_duration = max(r['duration'] for r in results) if results else 0
    
    safe_print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Auto-sync completed: "
               f"{success_count} success, {error_count} errors, "
               f"longest site took {total_max_duration}s")

if __name__ == '__main__':
    main()
