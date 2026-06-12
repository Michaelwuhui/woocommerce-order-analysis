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
from oid_utils import make_oid, site_id_for_source  # cross-site-safe surrogate order id

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
# 孤儿单清理保护阈值（P0-a 数据保护改造）
# 站点若被回滚到旧库，会突然出现大量"本地有、远程没有"的孤儿单；超过阈值即判定为
# 疑似数据事故，跳过删除并告警，避免本地备份随站点一起丢数据。
# -----------------------------
SUSPICIOUS_ORPHAN_ABS = 20      # 单站单次孤儿单绝对数量阈值
SUSPICIOUS_ORPHAN_PCT = 0.10    # 或孤儿单占该站本地订单的比例阈值
SUSPICIOUS_ORPHAN_PCT_MIN = 5   # 百分比规则的最小触发量（小站正常删几单不误报）

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


def fetch_all_remote_order_ids(wcapi, site_url):
    """从WooCommerce API获取所有订单ID列表"""
    all_ids = set()
    page = 1
    per_page = 100  # 使用较大的分页以减少请求次数
    max_retries = 3
    retry_count = 0
    
    print(f"正在获取 {site_url} 的所有远程订单ID...")
    
    while True:
        try:
            time.sleep(random.uniform(0.5, 1))
            
            # 只获取订单ID，减少数据传输
            response = wcapi.get("orders", params={
                "per_page": per_page,
                "page": page,
                "_fields": "id"  # 只返回ID字段
            })
            
            if response.status_code != 200:
                if retry_count < max_retries:
                    retry_count += 1
                    time.sleep(2)
                    continue
                else:
                    print(f"获取远程订单ID失败: HTTP {response.status_code}")
                    break
            
            data = response.json()
            if not data:
                break
            
            for order in data:
                all_ids.add(str(order['id']))
            
            print(f"已获取 {len(all_ids)} 个远程订单ID (页 {page})")
            
            page += 1
            retry_count = 0
            
        except Exception as e:
            print(f"获取远程订单ID时出错: {e}")
            if retry_count < max_retries:
                retry_count += 1
                time.sleep(2)
                continue
            else:
                break
    
    print(f"共获取 {len(all_ids)} 个远程订单ID")
    return all_ids


def ensure_orders_archive(connection):
    """确保归档表 orders_archive 存在，且列与 orders 对齐（外加 archived_at / archive_reason）。

    用于在物理删除孤儿单之前留底；对 orders 后续新增的列自动补齐，避免 schema 漂移导致归档失败。
    """
    cursor = connection.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS orders_archive AS SELECT * FROM orders WHERE 0")

    def _cols(table):
        return [row[1] for row in cursor.execute(f"PRAGMA table_info({table})").fetchall()]

    archive_cols = set(_cols('orders_archive'))
    for col in _cols('orders'):
        if col not in archive_cols:
            cursor.execute(f'ALTER TABLE orders_archive ADD COLUMN "{col}"')
            archive_cols.add(col)
    for meta_col in ('archived_at', 'archive_reason'):
        if meta_col not in archive_cols:
            cursor.execute(f"ALTER TABLE orders_archive ADD COLUMN {meta_col} TEXT")
    connection.commit()


def record_sync_alert(connection, site_url, alert_type, detail):
    """写入一条同步告警（如疑似站点回滚），供人工排查 / 后续在应用内展示。"""
    try:
        cursor = connection.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS sync_alerts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT,
                site_url TEXT,
                alert_type TEXT,
                detail TEXT,
                acknowledged INTEGER DEFAULT 0
            )
            """
        )
        cursor.execute(
            "INSERT INTO sync_alerts (created_at, site_url, alert_type, detail) VALUES (?, ?, ?, ?)",
            (datetime.now().isoformat(), site_url, alert_type, detail),
        )
        connection.commit()
    except Exception as e:
        print(f"记录同步告警失败: {e}")


def archive_orphaned_orders(site_url, remote_ids):
    """处理"本地有、远程没有"的孤儿订单（P0-a 数据保护，取代原先的直接物理删除）。

    - 小批量（正常的人工删单/去重）：先归档到 orders_archive，再从 orders 删除，对应用透明。
    - 大批量（疑似站点回滚/数据事故）：不删除、不归档，保留在 orders 中以免影响发货/对账，
      并写入 sync_alerts 告警等待人工确认——这是"站点崩了也不丢最新订单"的核心保护。

    返回从 orders 实际移除（已归档）的订单数。
    """
    local_ids = get_order_ids_from_db(site_url)

    if not local_ids:
        print(f"站点 {site_url} 本地无订单，跳过删除检查")
        return 0

    if not remote_ids:
        print(f"未获取到远程订单ID，跳过删除以避免误删")
        return 0

    # remote_ids 是裸 WC post id；本地 id 是跨站代理 id（"<sites.id>-<woo_id>"）。
    # 差集前必须把远程归一到同一代理空间，否则每个本地订单都会被判为孤儿。
    _conn_sid = create_database_connection()
    site_id = site_id_for_source(_conn_sid, site_url) if _conn_sid else None
    if _conn_sid:
        _conn_sid.close()
    if site_id is None:
        print(f"站点 {site_url} 不在 sites 表，跳过孤儿清理以免误删")
        return 0
    remote_ids = {make_oid(site_id, rid) for rid in remote_ids}

    local_set = set(local_ids)
    orphaned_ids = local_set - remote_ids

    if not orphaned_ids:
        print(f"站点 {site_url} 没有需要处理的孤立订单")
        return 0

    connection = create_database_connection()
    if not connection:
        return 0

    try:
        n_orphan = len(orphaned_ids)
        n_local = len(local_set)
        orphan_params = list(orphaned_ids)

        # —— 保护阈值：疑似站点回滚 / 数据事故 ——
        is_suspicious = (n_orphan >= SUSPICIOUS_ORPHAN_ABS) or (
            n_orphan >= SUSPICIOUS_ORPHAN_PCT_MIN and n_orphan >= SUSPICIOUS_ORPHAN_PCT * n_local
        )
        if is_suspicious:
            detail = (
                f"检测到 {n_orphan} 个孤儿订单（占本地 {n_local} 单的 {n_orphan / n_local:.0%}），"
                f"超过保护阈值，疑似站点回滚/数据事故；已跳过删除并保留这些订单。"
                f"示例ID: {sorted(orphaned_ids)[:10]}"
            )
            print("\n" + "!" * 72)
            print(f"⚠️  [数据保护] 站点 {site_url}: {detail}")
            print("!" * 72 + "\n")
            record_sync_alert(connection, site_url, 'suspected_rollback', detail)
            return 0

        # —— 正常小批量清理：先归档，再删除 ——
        ensure_orders_archive(connection)
        cursor = connection.cursor()
        order_cols = [row[1] for row in cursor.execute("PRAGMA table_info(orders)").fetchall()]
        col_list = ', '.join(f'"{c}"' for c in order_cols)
        placeholders = ','.join(['?'] * n_orphan)
        now_iso = datetime.now().isoformat()

        # 去重：同一订单若此前归档过，先删旧归档行，只保留最新一次
        cursor.execute(
            f"DELETE FROM orders_archive WHERE source = ? AND id IN ({placeholders})",
            [site_url] + orphan_params,
        )
        cursor.execute(
            f"INSERT INTO orders_archive ({col_list}, archived_at, archive_reason) "
            f"SELECT {col_list}, ?, ? FROM orders WHERE source = ? AND id IN ({placeholders})",
            [now_iso, 'orphaned_remote_deleted', site_url] + orphan_params,
        )
        archived = cursor.rowcount
        cursor.execute(
            f"DELETE FROM orders WHERE source = ? AND id IN ({placeholders})",
            [site_url] + orphan_params,
        )
        deleted_count = cursor.rowcount
        connection.commit()
        print(
            f"已归档并移除 {deleted_count} 个孤立订单（归档 {archived} 条）: "
            f"{sorted(orphaned_ids)[:10]}{'...' if n_orphan > 10 else ''}"
        )
        return deleted_count
    except Exception as e:
        print(f"处理孤立订单时出错: {e}")
        try:
            connection.rollback()
        except Exception:
            pass
        return 0
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

        # WC-managed columns. Local-only columns (is_undelivered, is_problem_return
        # and friends) are deliberately NOT listed — they're set by the web UI
        # and must survive every deep sync. INSERT OR REPLACE wipes ALL columns
        # not in the list because it DELETEs the row first; UPSERT only touches
        # the columns we name.
        wc_fields = [
            'id', 'parent_id', 'number', 'order_key', 'created_via', 'version', 'status', 'currency',
            'date_created', 'date_created_gmt', 'date_modified', 'date_modified_gmt',
            'discount_total', 'discount_tax', 'shipping_total', 'shipping_tax', 'cart_tax',
            'total', 'total_tax', 'prices_include_tax', 'customer_id', 'customer_ip_address',
            'customer_user_agent', 'customer_note', 'billing', 'shipping', 'payment_method',
            'payment_method_title', 'transaction_id', 'date_paid', 'date_paid_gmt',
            'date_completed', 'date_completed_gmt', 'cart_hash', 'meta_data', 'line_items',
            'tax_lines', 'shipping_lines', 'fee_lines', 'coupon_lines', 'refunds', 'set_paid', 'source'
        ]
        # woo_id keeps the raw per-site WC post id; id is the cross-site-safe
        # surrogate "<sites.id>-<woo_id>" (see oid_utils.py) so same-numbered
        # orders from different stores no longer collide under ON CONFLICT(id).
        all_columns = wc_fields + ['woo_id', 'updated_at']
        placeholders = ', '.join(['?'] * len(all_columns))
        update_set = ', '.join(f'{c} = excluded.{c}' for c in all_columns if c != 'id')
        insert_query = f"""
        INSERT INTO orders ({', '.join(all_columns)})
        VALUES ({placeholders})
        ON CONFLICT(id) DO UPDATE SET {update_set}
        """

        # 处理订单数据
        processed_orders = []
        for order in orders_data:
            woo_id = order.get('id')
            site_id = site_id_for_source(connection, order.get('source'))
            if site_id is None:
                # 未知站点来源 -> 无法构造安全的代理 id,跳过而非错误归并
                print(f"[save_orders_to_db] 跳过订单 {woo_id}: 未知来源 {order.get('source')!r}")
                continue
            oid = make_oid(site_id, woo_id)
            processed_order = []

            # 按照字段顺序处理数据
            fields = wc_fields

            for field in fields:
                value = order.get(field)

                # 特殊处理
                if field == 'id':
                    processed_order.append(oid)
                elif field == 'set_paid':
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

            # 添加 woo_id 与 updated_at 字段
            processed_order.append(woo_id)
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

def main(incremental=True, sync_status=True, start_date=None, clean_deleted=False):
    """主函数
    
    Args:
        incremental: 是否增量同步
        sync_status: 是否同步状态更新
        start_date: 起始日期
        clean_deleted: 是否清理已删除的订单
    """
    print("开始 WooCommerce 订单同步程序（SQLite版本）...")
    
    if clean_deleted:
        print("注意: 已启用清理同步——孤儿单先归档到 orders_archive 再删除；疑似回滚的大批量将跳过并告警")
    
    # 创建订单表
    create_orders_table()
    
    # 获取站点配置（优先从数据库，否则使用硬编码）
    sites = get_sites()
    
    total_deleted = 0
    
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
        # When called as a "deep sync" (incremental=False), we deliberately
        # leave last_order_date=None so the API fetches the entire history.
        
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
        
        # 清理已删除的订单
        if clean_deleted:
            print(f"\n开始检查站点 {site['url']} 的已删除订单...")
            remote_ids = fetch_all_remote_order_ids(wcapi, site['url'])
            deleted_count = archive_orphaned_orders(site['url'], remote_ids)
            total_deleted += deleted_count
        
        print(f"站点 {site['url']} 处理完成，等待处理下一个站点...")
        time.sleep(random.uniform(5, 10))
    
    print("\n所有站点处理完成！")
    if clean_deleted:
        print(f"共归档并清理 {total_deleted} 个孤立订单（疑似回滚的站点已跳过并告警）")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="WooCommerce订单同步工具")
    parser.add_argument("--start", type=str, help="YYYY-MM-DD 起始日期")
    parser.add_argument("--clean", action="store_true", help="清理已在WooCommerce中删除的订单")
    args = parser.parse_args()
    
    start_iso = None
    if args.start:
        try:
            dt = datetime.strptime(args.start, "%Y-%m-%d")
            start_iso = dt.strftime("%Y-%m-%dT00:00:00")
        except Exception:
            start_iso = None
    
    main(incremental=True, sync_status=False, start_date=start_iso, clean_deleted=args.clean)
