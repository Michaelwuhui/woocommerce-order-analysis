# SQLite版本的WooCommerce订单同步脚本
import json
import os
import time
import random
import sqlite3
from datetime import datetime, timedelta

import httpx
from httpx import HTTPError, TimeoutException
from woocommerce import API

# 添加代理配置（如果需要使用代理）
PROXY_CONFIG = {
    # "http": "http://your-proxy-server:port",
    # "https": "https://your-proxy-server:port"
}

def create_robust_wcapi(url, consumer_key, consumer_secret, proxy_config=None):
    """创建具有重试机制和更好错误处理的 WooCommerce API 客户端"""
    try:
        # 初始化 WooCommerce API 客户端
        wcapi = API(
            url=url,
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            version="wc/v3",
            timeout=30
        )
        
        return wcapi
    except Exception as e:
        print(f"创建 WooCommerce API 客户端时出错: {e}")
        return None

# -----------------------------
# 多站点配置
# -----------------------------
sites = [
    {
        "url": "https://www.buchmistrz.pl",
        "ck": "ck_75e405e4a60395d1b76aaebb1bf9cda39f53373a",
        "cs": "cs_7b4c64a2aa0681754d85a35100b70ddf562a33ca"
    },
    {
        "url": "https://www.strefajednorazowek.pl",
        "ck": "ck_5ce4c80f4b4abe045c2ba3bbe3ca31525db1b101",
        "cs": "cs_0016d8fd65de3474b16879281804b6fc055b34f1"
    },
]

# -----------------------------
# SQLite 数据库配置
# -----------------------------
DB_FILE = 'woocommerce_orders.db'

# -----------------------------
# 工具函数
# -----------------------------
def create_database_connection():
    """创建SQLite数据库连接"""
    try:
        connection = sqlite3.connect(DB_FILE)
        return connection
    except Exception as e:
        print(f"创建数据库连接时出错: {e}")
        return None

def create_orders_table():
    """创建订单表"""
    connection = create_database_connection()
    if not connection:
        return
    
    try:
        cursor = connection.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS orders (
            id TEXT PRIMARY KEY,
            parent_id TEXT,
            number TEXT,
            order_key TEXT,
            created_via TEXT,
            version TEXT,
            status TEXT,
            currency TEXT,
            date_created TEXT,
            date_created_gmt TEXT,
            date_modified TEXT,
            date_modified_gmt TEXT,
            discount_total REAL,
            discount_tax REAL,
            shipping_total REAL,
            shipping_tax REAL,
            cart_tax REAL,
            total REAL,
            total_tax REAL,
            prices_include_tax INTEGER,
            customer_id TEXT,
            customer_ip_address TEXT,
            customer_user_agent TEXT,
            customer_note TEXT,
            billing TEXT,
            shipping TEXT,
            payment_method TEXT,
            payment_method_title TEXT,
            transaction_id TEXT,
            date_paid TEXT,
            date_paid_gmt TEXT,
            date_completed TEXT,
            date_completed_gmt TEXT,
            cart_hash TEXT,
            meta_data TEXT,
            line_items TEXT,
            tax_lines TEXT,
            shipping_lines TEXT,
            fee_lines TEXT,
            coupon_lines TEXT,
            refunds TEXT,
            set_paid INTEGER,
            source TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
        cursor.execute(create_table_query)
        connection.commit()
        print("订单表创建成功或已存在")
    except Exception as e:
        print(f"创建订单表时出错: {e}")
    finally:
        if connection:
            connection.close()

def get_last_order_date_from_db(site_url):
    """从数据库获取指定站点的最新订单日期"""
    connection = create_database_connection()
    if not connection:
        return None
    
    try:
        cursor = connection.cursor()
        query = """
        SELECT MAX(date_created) 
        FROM orders 
        WHERE source = ?
        """
        cursor.execute(query, (site_url,))
        result = cursor.fetchone()
        return result[0] if result[0] else None
    except Exception as e:
        print(f"从数据库获取最新订单日期时出错: {e}")
        return None
    finally:
        if connection:
            connection.close()

def get_order_ids_from_db(site_url):
    """从数据库获取指定站点的所有订单ID"""
    connection = create_database_connection()
    if not connection:
        return []
    
    try:
        cursor = connection.cursor()
        query = "SELECT id FROM orders WHERE source = ?"
        cursor.execute(query, (site_url,))
        results = cursor.fetchall()
        return [str(row[0]) for row in results]
    except Exception as e:
        print(f"从数据库获取订单ID时出错: {e}")
        return []
    finally:
        if connection:
            connection.close()

def save_orders_to_db(orders_data):
    """将订单数据保存到SQLite数据库"""
    if not orders_data:
        print("没有订单数据可保存")
        return
    
    connection = create_database_connection()
    if not connection:
        return
    
    try:
        cursor = connection.cursor()
        
        # 准备插入语句
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
        
        # 处理订单数据
        processed_orders = []
        for order in orders_data:
            processed_order = []
            
            # 按照字段顺序处理数据
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
                
                # 特殊处理
                if field == 'set_paid':
                    if isinstance(value, dict):
                        processed_order.append(0)
                    elif value is None:
                        processed_order.append(0)
                    else:
                        processed_order.append(1 if value else 0)
                elif field == 'prices_include_tax':
                    processed_order.append(1 if value else 0)
                elif isinstance(value, (dict, list)):
                    processed_order.append(json.dumps(value, ensure_ascii=False))
                else:
                    processed_order.append(value)
            
            # 添加 updated_at 字段
            processed_order.append(datetime.now().isoformat())
            
            processed_orders.append(tuple(processed_order))
        
        # 批量插入数据
        cursor.executemany(insert_query, processed_orders)
        connection.commit()
        print(f"已保存 {len(processed_orders)} 个订单到SQLite数据库")
        
    except Exception as e:
        print(f"保存订单数据到数据库时出错: {e}")
    finally:
        if connection:
            connection.close()

def fetch_orders_incrementally(wcapi, site_url, last_order_date=None):
    """增量获取订单"""
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
        print(f"获取 {site_url} 站点 {last_order_date} 之后的订单...")
    else:
        print(f"获取 {site_url} 站点的所有订单...")
    
    while True:
        try:
            time.sleep(random.uniform(3, 6))
            
            response = wcapi.get("orders", params=params)
            
            if response.status_code == 403:
                print(f"站点 {site_url} 返回 403 错误，可能启用了安全防护")
                if retry_count < max_retries:
                    retry_count += 1
                    print(f"第 {retry_count} 次重试...")
                    time.sleep(random.uniform(15, 30))
                    continue
                else:
                    break
            elif response.status_code != 200:
                print(f"站点 {site_url} 返回 HTTP {response.status_code} 错误")
                if response.status_code >= 500 and retry_count < max_retries:
                    retry_count += 1
                    print(f"第 {retry_count} 次重试...")
                    time.sleep(random.uniform(10, 20))
                    continue
                else:
                    break
                
            data = response.json()
            if not data:
                break
                
            orders.extend(data)
            print(f"已获取 {len(data)} 个订单，当前页: {page}")
            page += 1
            params['page'] = page
            retry_count = 0
            
            time.sleep(random.uniform(2, 5))
            
        except Exception as e:
            print(f"获取站点 {site_url} 订单时发生错误: {e}")
            if retry_count < max_retries:
                retry_count += 1
                print(f"第 {retry_count} 次重试...")
                time.sleep(random.uniform(15, 30))
                continue
            else:
                break
    
    return orders

def main(incremental=True, sync_status=True):
    """主函数"""
    print("开始 WooCommerce 订单同步程序（SQLite版本）...")
    
    # 创建订单表
    create_orders_table()
    
    for site in sites:
        print(f"\n处理站点: {site['url']}")
        
        # 创建 WooCommerce API 客户端
        wcapi = create_robust_wcapi(site['url'], site['ck'], site['cs'], PROXY_CONFIG)
        if not wcapi:
            print(f"无法创建站点 {site['url']} 的 API 客户端，跳过...")
            continue
        
        # 获取最新订单日期（如果是增量同步）
        last_order_date = None
        if incremental:
            last_order_date = get_last_order_date_from_db(site['url'])
        
        # 获取订单数据
        orders = fetch_orders_incrementally(wcapi, site['url'], last_order_date)
        
        if orders:
            # 为每个订单添加来源站点信息
            for order in orders:
                order['source'] = site['url']
            
            print(f"从站点 {site['url']} 获取到 {len(orders)} 个订单")
            
            # 保存到数据库
            save_orders_to_db(orders)
        else:
            print(f"站点 {site['url']} 没有新订单")
        
        print(f"站点 {site['url']} 处理完成，等待处理下一个站点...")
        time.sleep(random.uniform(5, 10))
    
    print("\n所有站点处理完成！")

if __name__ == "__main__":
    # 运行主程序
    main(incremental=True, sync_status=False)