
import json
import time
import random
import sqlite3
from datetime import datetime
from woocommerce import API

# Database configuration
DB_FILE = 'woocommerce_orders.db'

# Proxy configuration (optional)
PROXY_CONFIG = {}

def create_robust_wcapi(url, consumer_key, consumer_secret, proxy_config=None):
    """Create robust WooCommerce API client"""
    try:
        wcapi = API(
            url=url,
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            version="wc/v3",
            timeout=30
        )
        return wcapi
    except Exception as e:
        print(f"Error creating WooCommerce API client: {e}")
        return None

def create_database_connection():
    """Create SQLite database connection"""
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

def save_orders_to_db(orders_data):
    """Save orders to SQLite database"""
    if not orders_data:
        return
    
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
        if connection:
            connection.close()

def fetch_orders_incrementally(wcapi, site_url, last_order_date=None, progress_callback=None):
    """Fetch orders incrementally"""
    orders = []
    page = 1
    per_page = 25
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
            time.sleep(random.uniform(0.5, 1)) 
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

            save_orders_to_db(data)
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

def fetch_orders_modified_after(wcapi, site_url, modified_after=None, progress_callback=None):
    """Fetch modified orders"""
    orders = []
    page = 1
    per_page = 25
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
            time.sleep(random.uniform(0.5, 1))
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
                
            save_orders_to_db(data)
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

def sync_site(url, consumer_key, consumer_secret, progress_callback=None):
    """Sync a single site"""
    if progress_callback: progress_callback(f"Connecting to {url}...")
    
    wcapi = create_robust_wcapi(url, consumer_key, consumer_secret, PROXY_CONFIG)
    if not wcapi:
        return {"status": "error", "message": "Failed to create API client"}
    
    try:
        # 1. Fetch new orders
        last_order_date = get_last_order_date_from_db(url)
        if progress_callback: 
            if last_order_date:
                progress_callback(f"Last sync date: {last_order_date}")
            else:
                progress_callback("First time sync (full history)...")
                
        new_orders = fetch_orders_incrementally(wcapi, url, last_order_date, progress_callback)
        
        # 2. Fetch updated orders
        last_modified = get_last_modified_date_from_db(url)
        updated_orders = fetch_orders_modified_after(wcapi, url, last_modified, progress_callback)
        
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
