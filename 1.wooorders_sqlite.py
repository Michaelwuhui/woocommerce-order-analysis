# SQLite版本的WooCommerce订单同步脚本
import json
import os
import time
import random
import sqlite3
from datetime import datetime, timedelta
import argparse

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
# 硬编码站点配置（作为备用/初始数据）
# Web 应用会优先使用数据库中的 sites 表配置
# -----------------------------
HARDCODED_SITES = [
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
    {
        "url": "https://www.vapicoau.com",
        "ck": "ck_72397759d2c45840107b9d9929ea3dbf4743c28a",
        "cs": "cs_555a5c961438a761f953c624a76de60c243d4a97"
    },
    {
        "url": "https://www.vapeprime.pl",
        "ck": "ck_b4d828f212b61228ae94506f6bdb71b5a4f41e8c",
        "cs": "cs_06af71df5c384c910726d23bcb2d34f94f783253"
    },
    {
        "url": "https://www.vapeprimeau.com",
        "ck": "ck_e8ff54af41a9de304614bb504665e1de53162da5",
        "cs": "cs_77a37538dd6c480227a1f57e28b2903411220a6c"
    },
    {
        "url": "https://www.vaportrail.ae",
        "ck": "ck_646ed2e068e58bda88e5af7ee43c36f981087416",
        "cs": "cs_0cf22f59596be752c3243d80294e003bdb87bd2f"
    },
    {
        "url": "https://vaporburst.ae",
        "ck": "cs_838aa96f94e0cabc041be0aeaa541804871e16a9",
        "cs": "ck_2fb0cf7665a180445bfb28fbdbd7f66281bbf1c7"
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


def get_sites_from_db():
    """从数据库 sites 表读取站点配置"""
    try:
        connection = create_database_connection()
        if not connection:
            return []
        cursor = connection.cursor()
        cursor.execute('SELECT url, consumer_key, consumer_secret FROM sites')
        rows = cursor.fetchall()
        connection.close()
        
        if rows:
            sites = [{'url': r[0], 'ck': r[1], 'cs': r[2]} for r in rows]
            print(f"从数据库加载了 {len(sites)} 个站点配置")
            return sites
        return []
    except Exception as e:
        print(f"从数据库读取站点配置失败: {e}")
        return []


def get_sites():
    """获取站点配置，优先从数据库读取，若无数据则使用硬编码配置"""
    db_sites = get_sites_from_db()
    if db_sites:
        return db_sites
    print("数据库无站点配置，使用硬编码配置")
    return HARDCODED_SITES

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

def get_last_modified_date_from_db(site_url):
    connection = create_database_connection()
    if not connection:
        return None
    try:
        cursor = connection.cursor()
        query = """
        SELECT MAX(date_modified)
        FROM orders
        WHERE source = ?
        """
        cursor.execute(query, (site_url,))
        result = cursor.fetchone()
        return result[0] if result[0] else None
    except Exception as e:
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
    """增量获取订单（逐页保存）"""
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
            time.sleep(random.uniform(1, 2))

            response = wcapi.get("orders", params=params)

            if response.status_code == 403:
                print(f"站点 {site_url} 返回 403 错误，可能启用了安全防护")
                if retry_count < max_retries:
                    retry_count += 1
                    print(f"第 {retry_count} 次重试...")
                    time.sleep(random.uniform(5, 10))
                    continue
                else:
                    break
            elif response.status_code != 200:
                print(f"站点 {site_url} 返回 HTTP {response.status_code} 错误")
                if response.status_code >= 500 and retry_count < max_retries:
                    retry_count += 1
                    print(f"第 {retry_count} 次重试...")
                    time.sleep(random.uniform(5, 10))
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
            print(f"已获取 {len(data)} 个订单，当前页: {page}")
            page += 1
            params['page'] = page
            retry_count = 0

            time.sleep(random.uniform(1, 2))

        except Exception as e:
            print(f"获取站点 {site_url} 订单时发生错误: {e}")
            if retry_count < max_retries:
                retry_count += 1
                print(f"第 {retry_count} 次重试...")
                time.sleep(random.uniform(5, 10))
                continue
            else:
                break

    return orders

def fetch_orders_modified_after(wcapi, site_url, modified_after=None):
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
        print(f"获取 {site_url} 站点 {modified_after} 之后修改的订单...")
    else:
        print(f"获取 {site_url} 站点的近期修改订单...")
    while True:
        try:
            time.sleep(random.uniform(1, 2))
            response = wcapi.get("orders", params=params)
            if response.status_code == 403:
                if retry_count < max_retries:
                    retry_count += 1
                    time.sleep(random.uniform(5, 10))
                    continue
                else:
                    break
            elif response.status_code != 200:
                if response.status_code >= 500 and retry_count < max_retries:
                    retry_count += 1
                    time.sleep(random.uniform(5, 10))
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
            print(f"已获取修改 {len(data)} 个订单，当前页: {page}")
            page += 1
            params['page'] = page
            retry_count = 0
            time.sleep(random.uniform(1, 2))
        except Exception:
            if retry_count < max_retries:
                retry_count += 1
                time.sleep(random.uniform(5, 10))
                continue
            else:
                break
    return orders

def main(incremental=True, sync_status=True, start_date=None):
    """主函数"""
    print("开始 WooCommerce 订单同步程序（SQLite版本）...")
    
    # 创建订单表
    create_orders_table()
    
    # 获取站点配置（优先从数据库，否则使用硬编码）
    sites = get_sites()
    
    for site in sites:
        print(f"\n处理站点: {site['url']}")
        
        # 创建 WooCommerce API 客户端
        wcapi = create_robust_wcapi(site['url'], site['ck'], site['cs'], PROXY_CONFIG)
        if not wcapi:
            print(f"无法创建站点 {site['url']} 的 API 客户端，跳过...")
            continue
        
        last_order_date = None
        if start_date:
            last_order_date = start_date
        elif incremental:
            last_order_date = get_last_order_date_from_db(site['url'])
        
        # 获取订单数据
        orders = fetch_orders_incrementally(wcapi, site['url'], last_order_date)
        
        if orders:
            print(f"从站点 {site['url']} 获取到 {len(orders)} 个订单")
        else:
            print(f"站点 {site['url']} 没有新订单")

        modified_after = start_date or get_last_modified_date_from_db(site['url'])
        updated_orders = fetch_orders_modified_after(wcapi, site['url'], modified_after)
        if updated_orders:
            print(f"站点 {site['url']} 获取到 {len(updated_orders)} 个修改更新的订单")
        
        print(f"站点 {site['url']} 处理完成，等待处理下一个站点...")
        time.sleep(random.uniform(5, 10))
    
    print("\n所有站点处理完成！")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", type=str, help="YYYY-MM-DD 起始日期")
    args = parser.parse_args()
    start_iso = None
    if args.start:
        try:
            dt = datetime.strptime(args.start, "%Y-%m-%d")
            start_iso = dt.strftime("%Y-%m-%dT00:00:00")
        except Exception:
            start_iso = None
    main(incremental=True, sync_status=False, start_date=start_iso)
