#!/usr/bin/env python3
"""
Auto-sync script for cron execution.
This script is called by cron to automatically sync all sites.
"""
import sqlite3
import json
from datetime import datetime
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import sync_utils

DB_FILE = 'woocommerce_orders.db'

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

def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Auto-sync started")
    
    conn = get_db_connection()
    sites = conn.execute('SELECT * FROM sites').fetchall()
    conn.close()
    
    if not sites:
        print("No sites configured")
        return
    
    for site in sites:
        site_url = site['url']
        print(f"Syncing: {site_url}")
        
        start_time = datetime.now()
        
        def progress_callback(msg):
            print(f"  {msg}")
        
        try:
            result = sync_utils.sync_site(
                site['url'],
                site['consumer_key'],
                site['consumer_secret'],
                progress_callback
            )
            
            duration = int((datetime.now() - start_time).total_seconds())
            
            if result['status'] == 'success':
                # Update last_sync time
                conn = get_db_connection()
                conn.execute('UPDATE sites SET last_sync = ? WHERE id = ?',
                            (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), site['id']))
                conn.commit()
                conn.close()
                
                log_sync(
                    site['id'], site_url, 'success',
                    f"Synced successfully: {result.get('new_orders', 0)} new, {result.get('updated_orders', 0)} updated",
                    result.get('new_orders', 0), result.get('updated_orders', 0), duration
                )
                print(f"  ✓ Success ({duration}s)")
            else:
                log_sync(site['id'], site_url, 'error', result.get('message', 'Unknown error'), 0, 0, duration)
                print(f"  ✗ Error: {result.get('message', 'Unknown error')}")
                
        except Exception as e:
            duration = int((datetime.now() - start_time).total_seconds())
            log_sync(site['id'], site_url, 'error', str(e), 0, 0, duration)
            print(f"  ✗ Exception: {e}")
    
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Auto-sync completed")

if __name__ == '__main__':
    main()
