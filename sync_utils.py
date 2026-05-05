import json
import time
import sqlite3
import threading
import concurrent.futures
from datetime import datetime
from woocommerce import API

# Database configuration
DB_FILE = 'woocommerce_orders.db'

# Proxy configuration (optional)
PROXY_CONFIG = {}

# 线程局部存储，用于数据库连接复用
_thread_local = threading.local()

def get_thread_db_connection():
    """获取当前线程的数据库连接（复用，启用 WAL 模式支持并发）"""
    if not hasattr(_thread_local, 'connection') or _thread_local.connection is None:
        conn = sqlite3.connect(DB_FILE, timeout=30)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=30000")
        _thread_local.connection = conn
    return _thread_local.connection

def close_thread_db_connection():
    """关闭当前线程的数据库连接"""
    if hasattr(_thread_local, 'connection') and _thread_local.connection is not None:
        try:
            _thread_local.connection.close()
        except:
            pass
        _thread_local.connection = None

def create_robust_wcapi(url, consumer_key, consumer_secret, proxy_config=None):
    """Create robust WooCommerce API client"""
    try:
        wcapi = API(
            url=url,
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            version="wc/v3",
            timeout=60
        )
        return wcapi
    except Exception as e:
        print(f"Error creating WooCommerce API client: {e}")
        return None

def create_database_connection():
    """Create SQLite database connection (兼容旧接口，新代码建议使用 get_thread_db_connection)"""
    try:
        connection = sqlite3.connect(DB_FILE)
        return connection
    except Exception as e:
        print(f"Error creating database connection: {e}")
        return None

def get_last_order_date_from_db(site_url):
    """Get last order date for a site from DB"""
    connection = create_database_connection()
    if not connection:
        return None
    
    try:
        cursor = connection.cursor()
        query = "SELECT MAX(date_created) FROM orders WHERE source = ?"
        cursor.execute(query, (site_url,))
        result = cursor.fetchone()
        return result[0] if result[0] else None
    except Exception as e:
        print(f"Error getting last order date: {e}")
        return None
    finally:
        if connection:
            connection.close()

def get_last_modified_date_from_db(site_url):
    """Get last modified date for a site from DB"""
    connection = create_database_connection()
    if not connection:
        return None
    try:
        cursor = connection.cursor()
        query = "SELECT MAX(date_modified) FROM orders WHERE source = ?"
        cursor.execute(query, (site_url,))
        result = cursor.fetchone()
        return result[0] if result[0] else None
    except Exception as e:
        return None
    finally:
        if connection:
            connection.close()

def save_orders_to_db(orders_data, connection=None):
    """Save orders to SQLite database"""
    if not orders_data:
        return
    
    own_connection = connection is None
    if own_connection:
        connection = create_database_connection()
        if not connection:
            return
    
    try:
        cursor = connection.cursor()
        
        insert_query = """
        INSERT OR REPLACE INTO orders (
            id, parent_id, number, order_key, created_via, version, status, currency,
            date_created, date_created_gmt, date_modified, date_modified_gmt,
            discount_total, discount_tax, shipping_total, shipping_tax, cart_tax,
            total, total_tax, prices_include_tax, customer_id, customer_ip_address,
            customer_user_agent, customer_note, billing, shipping, payment_method,
            payment_method_title, transaction_id, date_paid, date_paid_gmt,
            date_completed, date_completed_gmt, cart_hash, meta_data, line_items,
            tax_lines, shipping_lines, fee_lines, coupon_lines, refunds, set_paid, source,
            updated_at
        ) VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
        """
        
        # Filter out checkout-draft orders - they should not be synced
        orders_data = [o for o in orders_data if o.get('status') != 'checkout-draft']
        
        if not orders_data:
            return
        
        processed_orders = []
        for order in orders_data:
            processed_order = []
            fields = [
                'id', 'parent_id', 'number', 'order_key', 'created_via', 'version', 'status', 'currency',
                'date_created', 'date_created_gmt', 'date_modified', 'date_modified_gmt',
                'discount_total', 'discount_tax', 'shipping_total', 'shipping_tax', 'cart_tax',
                'total', 'total_tax', 'prices_include_tax', 'customer_id', 'customer_ip_address',
                'customer_user_agent', 'customer_note', 'billing', 'shipping', 'payment_method',
                'payment_method_title', 'transaction_id', 'date_paid', 'date_paid_gmt',
                'date_completed', 'date_completed_gmt', 'cart_hash', 'meta_data', 'line_items',
                'tax_lines', 'shipping_lines', 'fee_lines', 'coupon_lines', 'refunds', 'set_paid', 'source'
            ]
            
            for field in fields:
                value = order.get(field)
                if field == 'set_paid':
                    if isinstance(value, dict) or value is None:
                        processed_order.append(0)
                    else:
                        processed_order.append(1 if value else 0)
                elif field == 'prices_include_tax':
                    processed_order.append(1 if value else 0)
                elif isinstance(value, (dict, list)):
                    processed_order.append(json.dumps(value, ensure_ascii=False))
                else:
                    processed_order.append(value)
            
            processed_order.append(datetime.now().isoformat())
            processed_orders.append(tuple(processed_order))
        
        cursor.executemany(insert_query, processed_orders)
        connection.commit()
        
    except Exception as e:
        print(f"Error saving orders: {e}")
    finally:
        if own_connection and connection:
            connection.close()

def fetch_orders_incrementally(wcapi, site_url, last_order_date=None, progress_callback=None, connection=None):
    """Fetch orders incrementally"""
    orders = []
    page = 1
    per_page = 100
    max_retries = 3
    retry_count = 0

    params = {
        "per_page": per_page,
        "page": page,
        "expand": "line_items,shipping_lines,tax_lines,fee_lines,coupon_lines,refunds"
    }

    if last_order_date:
        params['after'] = last_order_date
        if progress_callback: progress_callback(f"Fetching orders after {last_order_date}...")

    while True:
        try:
            if progress_callback: progress_callback(f"Fetching page {page}...")
            response = wcapi.get("orders", params=params)

            if response.status_code != 200:
                if retry_count < max_retries:
                    retry_count += 1
                    if progress_callback: progress_callback(f"Error {response.status_code}, retrying ({retry_count}/{max_retries})...")
                    time.sleep(2)
                    continue
                else:
                    if progress_callback: progress_callback(f"Failed after max retries.")
                    break

            data = response.json()
            if not data:
                if progress_callback: progress_callback(f"No more orders found.")
                break

            for order in data:
                order['source'] = site_url

            save_orders_to_db(data, connection=connection)
            orders.extend(data)
            
            if progress_callback: progress_callback(f"Saved {len(data)} orders from page {page}.")
            
            page += 1
            params['page'] = page
            retry_count = 0

        except Exception as e:
            if progress_callback: progress_callback(f"Error: {str(e)}")
            if retry_count < max_retries:
                retry_count += 1
                time.sleep(2)
                continue
            else:
                break

    return orders

def fetch_orders_modified_after(wcapi, site_url, modified_after=None, progress_callback=None, connection=None):
    """Fetch modified orders"""
    orders = []
    page = 1
    per_page = 100
    max_retries = 3
    retry_count = 0
    
    params = {
        "per_page": per_page,
        "page": page,
        "expand": "line_items,shipping_lines,tax_lines,fee_lines,coupon_lines,refunds"
    }
    
    if modified_after:
        params['modified_after'] = modified_after
        if progress_callback: progress_callback(f"Checking for updates after {modified_after}...")
        
    while True:
        try:
            if progress_callback: progress_callback(f"Fetching updates page {page}...")
            response = wcapi.get("orders", params=params)
            
            if response.status_code != 200:
                if retry_count < max_retries:
                    retry_count += 1
                    time.sleep(2)
                    continue
                else:
                    break
                    
            data = response.json()
            if not data:
                break
                
            for order in data:
                order['source'] = site_url
                
            save_orders_to_db(data, connection=connection)
            orders.extend(data)
            
            if progress_callback: progress_callback(f"Updated {len(data)} orders from page {page}.")
            
            page += 1
            params['page'] = page
            retry_count = 0
            
        except Exception as e:
            if progress_callback: progress_callback(f"Error: {str(e)}")
            if retry_count < max_retries:
                retry_count += 1
                time.sleep(2)
                continue
            else:
                break
                
    return orders

def sync_order_notes(wcapi, site_url, connection=None):
    """Fetch and sync order notes for active orders.

    The WC REST API requires the internal post ID (orders.id), not the
    customer-facing order number. Sites that use the Sequential Order Numbers
    plugin (vapeprimeau.com, vapego.pl, …) have number != id, so calling
    /orders/{number}/notes returns 404. Always use orders.id for the API call.
    """
    if not connection:
        return

    try:
        cursor = connection.cursor()

        # Use the WC internal post id (orders.id) for the API call. Number can
        # be a customer-facing sequential number that the API does not accept.
        cursor.execute('''
            SELECT id
            FROM orders
            WHERE source = ? AND status IN ('processing', 'offline', 'on-hold')
        ''', (site_url,))

        active_order_ids = [row[0] for row in cursor.fetchall()]

        if not active_order_ids:
            return

        def fetch_notes_for_order(order_id):
            try:
                response = wcapi.get(f"orders/{order_id}/notes")
                if response.status_code == 200:
                    notes_data = response.json()
                    for note in notes_data:
                        note['_local_order_id'] = order_id
                    return notes_data
            except Exception as e:
                print(f"Error fetching notes for order {order_id}: {e}")
            return []

        all_notes = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future_to_order = {executor.submit(fetch_notes_for_order, oid): oid for oid in active_order_ids}

            for future in concurrent.futures.as_completed(future_to_order):
                order_id = future_to_order[future]
                try:
                    notes = future.result()
                    if notes:
                        all_notes.extend(notes)
                except Exception as exc:
                    print(f'{order_id} generated an exception: {exc}')

        if all_notes:
            # Dedupe by (order_id, wc_note_id). WC note IDs are per-site auto-increment
            # so they collide across sites — using them as the local PK silently
            # overwrites notes from earlier-synced sites.
            insert_query = """
            INSERT OR REPLACE INTO order_notes (
                wc_note_id, order_id, note, date_created, customer_note, author, added_by_user
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """

            processed_notes = []
            for note in all_notes:
                local_order_id = note['_local_order_id']
                added_by_user = 1 if note.get('added_by_user', False) else 0
                customer_note = 1 if note.get('customer_note', False) else 0

                processed_notes.append((
                    note.get('id'),
                    local_order_id,
                    note.get('note', ''),
                    note.get('date_created', ''),
                    customer_note,
                    note.get('author', ''),
                    added_by_user
                ))

            if processed_notes:
                cursor.executemany(insert_query, processed_notes)
                connection.commit()

    except Exception as e:
        print(f"Error syncing order notes: {e}")

def sync_site(url, consumer_key, consumer_secret, progress_callback=None, sync_days=7, full_history=False):
    """Sync a single site

    Args:
        sync_days: Only sync orders modified in the last N days.
                   Set to 0 or None for sync from last known order date.
        full_history: When True, ignore all local cutoffs and fetch every order
                      page from the WooCommerce API. Use this for first-time
                      sync of a site or when local DB is missing historical data.
    """
    if progress_callback: progress_callback(f"Connecting to {url}...")

    wcapi = create_robust_wcapi(url, consumer_key, consumer_secret, PROXY_CONFIG)
    if not wcapi:
        return {"status": "error", "message": "Failed to create API client"}

    # 使用线程局部数据库连接，整个同步过程复用
    conn = get_thread_db_connection()

    try:
        # Calculate time window for modified_after
        modified_after = None
        if full_history:
            # No cutoff at all — fetch every page
            modified_after = None
            if progress_callback:
                progress_callback("Full history sync (no date filter)...")
        elif sync_days and sync_days > 0:
            from datetime import timedelta
            cutoff_date = datetime.now() - timedelta(days=sync_days)
            modified_after = cutoff_date.strftime("%Y-%m-%dT00:00:00")
            if progress_callback:
                progress_callback(f"Syncing orders modified in last {sync_days} days...")
        else:
            # Full sync: use last modified date from DB
            modified_after = get_last_modified_date_from_db(url)
            if progress_callback:
                if modified_after:
                    progress_callback(f"Full sync from {modified_after}...")
                else:
                    progress_callback("Full sync (all history)...")

        # 1. Fetch new orders (only for first-time sync or when no cutoff)
        new_orders = []
        if full_history or not sync_days or sync_days <= 0:
            last_order_date = None if full_history else get_last_order_date_from_db(url)
            if progress_callback:
                if last_order_date:
                    progress_callback(f"Fetching new orders after {last_order_date}...")
                else:
                    progress_callback("First time sync (full history)...")
            new_orders = fetch_orders_incrementally(wcapi, url, last_order_date, progress_callback, connection=conn)
        
        # 2. Fetch updated orders (within time window)
        updated_orders = fetch_orders_modified_after(wcapi, url, modified_after, progress_callback, connection=conn)
        
        # 3. Sync order notes for active orders
        if progress_callback: progress_callback("Syncing order notes...")
        sync_order_notes(wcapi, url, connection=conn)
        
        msg = f"Sync complete. New: {len(new_orders)}, Updated: {len(updated_orders)}"
        if progress_callback: progress_callback(msg)
        
        return {
            "status": "success", 
            "new_orders": len(new_orders), 
            "updated_orders": len(updated_orders)
        }
    except Exception as e:
        if progress_callback: progress_callback(f"Critical Error: {str(e)}")
        return {"status": "error", "message": str(e)}
    finally:
        close_thread_db_connection()

