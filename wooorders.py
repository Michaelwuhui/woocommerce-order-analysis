# mysql-connector-python
import json
import os
import time
import random
from datetime import datetime, timedelta

import mysql.connector
import httpx
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from mysql.connector import Error
from httpx import HTTPError, TimeoutException
from woocommerce import API

# 添加代理配置（如果需要使用代理）
# 注意：只有在确实有代理服务器时才取消注释下面的行
PROXY_CONFIG = {
    # "http": "http://your-proxy-server:port",
    # "https": "https://your-proxy-server:port"
    # 示例：
    # "http": "http://127.0.0.1:8080",
    # "https": "https://127.0.0.1:8080"
}

def create_robust_wcapi(url, consumer_key, consumer_secret, proxy_config=None):
    """创建具有重试机制和更好错误处理的 WooCommerce API 客户端"""
    try:
        # 配置重试策略
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        
        # 创建会话
        session = requests.Session()
        session.mount("http://", HTTPAdapter(max_retries=retry_strategy))
        session.mount("https://", HTTPAdapter(max_retries=retry_strategy))
        
        # 设置代理（如果提供且不为空）
        if proxy_config and proxy_config.get("http") and proxy_config.get("https"):
            session.proxies.update(proxy_config)
        
        # 设置更真实的请求头（基于我们成功的测试）
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/json",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "Cache-Control": "no-cache",
            "Referer": url + "/",
            "Origin": url
        })
        
        # 初始化 WooCommerce API 客户端
        wcapi = API(
            url=url,
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            version="wc/v3",
            timeout=30,
            requests_kwargs={
                "session": session
            }
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
    # 可以继续添加更多站点
]

# -----------------------------
# MySQL 数据库配置
# -----------------------------
db_config = {
    'host': 'localhost',
    'database': 'woocommerce_orders',
    'user': 'root',
    'password': '1207wlp@'
}

# -----------------------------
# 工具函数
# -----------------------------
def create_database_connection():
    """创建数据库连接"""
    try:
        connection = mysql.connector.connect(**db_config)
        if connection.is_connected():
            return connection
    except Error as e:
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
            id VARCHAR(50) PRIMARY KEY,
            parent_id VARCHAR(50),
            number VARCHAR(50),
            order_key VARCHAR(100),
            created_via VARCHAR(50),
            version VARCHAR(20),
            status VARCHAR(50),
            currency VARCHAR(10),
            date_created DATETIME,
            date_created_gmt DATETIME,
            date_modified DATETIME,
            date_modified_gmt DATETIME,
            discount_total DECIMAL(10, 2),
            discount_tax DECIMAL(10, 2),
            shipping_total DECIMAL(10, 2),
            shipping_tax DECIMAL(10, 2),
            cart_tax DECIMAL(10, 2),
            total DECIMAL(10, 2),
            total_tax DECIMAL(10, 2),
            prices_include_tax BOOLEAN,
            customer_id VARCHAR(50),
            customer_ip_address VARCHAR(50),
            customer_user_agent TEXT,
            customer_note TEXT,
            billing TEXT,
            shipping TEXT,
            payment_method VARCHAR(50),
            payment_method_title VARCHAR(100),
            transaction_id VARCHAR(100),
            date_paid DATETIME,
            date_paid_gmt DATETIME,
            date_completed DATETIME,
            date_completed_gmt DATETIME,
            cart_hash VARCHAR(100),
            meta_data TEXT,
            line_items TEXT,
            tax_lines TEXT,
            shipping_lines TEXT,
            fee_lines TEXT,
            coupon_lines TEXT,
            refunds TEXT,
            set_paid BOOLEAN,
            source VARCHAR(255),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
        )
        """
        cursor.execute(create_table_query)
        connection.commit()
        print("订单表创建成功或已存在")
    except Error as e:
        print(f"创建订单表时出错: {e}")
    finally:
        if connection.is_connected():
            cursor.close()
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
        WHERE source = %s
        """
        cursor.execute(query, (site_url,))
        result = cursor.fetchone()
        return result[0].strftime('%Y-%m-%dT%H:%M:%S') if result[0] else None
    except Error as e:
        print(f"从数据库获取最新订单日期时出错: {e}")
        return None
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

def get_order_ids_from_db(site_url):
    """从数据库获取指定站点的所有订单ID"""
    connection = create_database_connection()
    if not connection:
        return []
    
    try:
        cursor = connection.cursor()
        query = "SELECT id FROM orders WHERE source = %s"
        cursor.execute(query, (site_url,))
        results = cursor.fetchall()
        return [str(row[0]) for row in results]
    except Error as e:
        print(f"从数据库获取订单ID时出错: {e}")
        return []
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

def update_order_status_in_db(order_id, new_status, date_modified):
    """更新数据库中的订单状态和修改时间"""
    connection = create_database_connection()
    if not connection:
        return False
    
    try:
        cursor = connection.cursor()
        query = """
        UPDATE orders 
        SET status = %s, date_modified = %s, updated_at = CURRENT_TIMESTAMP
        WHERE id = %s
        """
        # 将 ISO 8601 格式转换为 MySQL datetime 格式
        if date_modified:
            try:
                dt = datetime.fromisoformat(date_modified.replace('Z', '+00:00'))
                mysql_date = dt.strftime('%Y-%m-%d %H:%M:%S')
            except:
                mysql_date = None
        else:
            mysql_date = None
            
        cursor.execute(query, (new_status, mysql_date, order_id))
        connection.commit()
        if cursor.rowcount > 0:
            print(f"数据库中订单 #{order_id} 状态已更新为 {new_status}")
            return True
        else:
            print(f"数据库中未找到订单 #{order_id}")
            return False
    except Error as e:
        print(f"更新数据库中订单 #{order_id} 状态时出错: {e}")
        return False
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

def save_orders_to_db(orders_data):
    """将订单数据保存到数据库"""
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
        INSERT INTO orders (
            id, parent_id, number, order_key, created_via, version, status, currency,
            date_created, date_created_gmt, date_modified, date_modified_gmt,
            discount_total, discount_tax, shipping_total, shipping_tax, cart_tax,
            total, total_tax, prices_include_tax, customer_id, customer_ip_address,
            customer_user_agent, customer_note, billing, shipping, payment_method,
            payment_method_title, transaction_id, date_paid, date_paid_gmt,
            date_completed, date_completed_gmt, cart_hash, meta_data, line_items,
            tax_lines, shipping_lines, fee_lines, coupon_lines, refunds, set_paid, source
        ) VALUES (
            %(id)s, %(parent_id)s, %(number)s, %(order_key)s, %(created_via)s, %(version)s,
            %(status)s, %(currency)s, %(date_created)s, %(date_created_gmt)s, %(date_modified)s,
            %(date_modified_gmt)s, %(discount_total)s, %(discount_tax)s, %(shipping_total)s,
            %(shipping_tax)s, %(cart_tax)s, %(total)s, %(total_tax)s, %(prices_include_tax)s,
            %(customer_id)s, %(customer_ip_address)s, %(customer_user_agent)s, %(customer_note)s,
            %(billing)s, %(shipping)s, %(payment_method)s, %(payment_method_title)s,
            %(transaction_id)s, %(date_paid)s, %(date_paid_gmt)s, %(date_completed)s,
            %(date_completed_gmt)s, %(cart_hash)s, %(meta_data)s, %(line_items)s,
            %(tax_lines)s, %(shipping_lines)s, %(fee_lines)s, %(coupon_lines)s,
            %(refunds)s, %(set_paid)s, %(source)s
        ) ON DUPLICATE KEY UPDATE
            parent_id = VALUES(parent_id),
            number = VALUES(number),
            order_key = VALUES(order_key),
            created_via = VALUES(created_via),
            version = VALUES(version),
            status = VALUES(status),
            currency = VALUES(currency),
            date_created = VALUES(date_created),
            date_created_gmt = VALUES(date_created_gmt),
            date_modified = VALUES(date_modified),
            date_modified_gmt = VALUES(date_modified_gmt),
            discount_total = VALUES(discount_total),
            discount_tax = VALUES(discount_tax),
            shipping_total = VALUES(shipping_total),
            shipping_tax = VALUES(shipping_tax),
            cart_tax = VALUES(cart_tax),
            total = VALUES(total),
            total_tax = VALUES(total_tax),
            prices_include_tax = VALUES(prices_include_tax),
            customer_id = VALUES(customer_id),
            customer_ip_address = VALUES(customer_ip_address),
            customer_user_agent = VALUES(customer_user_agent),
            customer_note = VALUES(customer_note),
            billing = VALUES(billing),
            shipping = VALUES(shipping),
            payment_method = VALUES(payment_method),
            payment_method_title = VALUES(payment_method_title),
            transaction_id = VALUES(transaction_id),
            date_paid = VALUES(date_paid),
            date_paid_gmt = VALUES(date_paid_gmt),
            date_completed = VALUES(date_completed),
            date_completed_gmt = VALUES(date_completed_gmt),
            cart_hash = VALUES(cart_hash),
            meta_data = VALUES(meta_data),
            line_items = VALUES(line_items),
            tax_lines = VALUES(tax_lines),
            shipping_lines = VALUES(shipping_lines),
            fee_lines = VALUES(fee_lines),
            coupon_lines = VALUES(coupon_lines),
            refunds = VALUES(refunds),
            set_paid = VALUES(set_paid),
            source = VALUES(source),
            updated_at = CURRENT_TIMESTAMP
        """
        
        # 处理订单数据
        processed_orders = []
        for order in orders_data:
            processed_order = {}
            for key, value in order.items():
                # 特殊处理 set_paid 字段，将其转换为布尔值
                if key == 'set_paid':
                    # 如果 set_paid 是字典类型（包含 href），则设置为 False
                    # 如果是布尔值，则保持原样
                    # 如果是 None，则设置为 False
                    if isinstance(value, dict):
                        processed_order[key] = False
                    elif value is None:
                        processed_order[key] = False
                    else:
                        processed_order[key] = bool(value)
                elif isinstance(value, dict) or isinstance(value, list):
                    processed_order[key] = json.dumps(value, ensure_ascii=False)
                else:
                    processed_order[key] = value
                    
                # 特殊处理日期字段
                if key in ['date_created', 'date_created_gmt', 'date_modified', 'date_modified_gmt', 
                          'date_paid', 'date_paid_gmt', 'date_completed', 'date_completed_gmt']:
                    if value:
                        try:
                            # 将 ISO 8601 格式转换为 MySQL datetime 格式
                            dt = datetime.fromisoformat(value.replace('Z', '+00:00'))
                            processed_order[key] = dt.strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            processed_order[key] = None
            
            # 确保 set_paid 字段存在，即使原始数据中没有
            if 'set_paid' not in processed_order:
                processed_order['set_paid'] = False
                
            processed_orders.append(processed_order)
        
        # 批量插入数据
        print("准备插入的数据示例（第一个订单）:")
        if processed_orders:
            for key, value in processed_orders[0].items():
                print(f"  {key}: {value} ({type(value)})")
        
        cursor.executemany(insert_query, processed_orders)
        connection.commit()
        print(f"已保存 {cursor.rowcount} 个订单到数据库")
        
    except Error as e:
        print(f"保存订单数据到数据库时出错: {e}")
        print(f"MySQL错误代码: {e.errno}")
        print(f"MySQL错误信息: {e.msg}")
        # 如果有额外的诊断信息
        if hasattr(e, 'sqlstate'):
            print(f"SQL状态: {e.sqlstate}")
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

# -----------------------------
# 主功能函数
# -----------------------------
def fetch_orders_incrementally(wcapi, site_url, last_order_date=None):
    """增量获取订单（获取完整字段）"""
    orders = []
    page = 1
    per_page = 25  # 进一步减少每页数量
    max_retries = 3  # 减少重试次数以避免过度重试
    retry_count = 0
    
    # 构建参数，请求完整的订单数据
    params = {
        "per_page": per_page, 
        "page": page,
        "expand": "line_items,shipping_lines,tax_lines,fee_lines,coupon_lines,refunds"  # 展开所有相关数据
    }
    
    # 如果有上次订单日期，则只获取该日期之后的订单
    if last_order_date:
        params['after'] = last_order_date
        print(f"获取 {site_url} 站点 {last_order_date} 之后的订单...")
    else:
        print(f"获取 {site_url} 站点的所有订单...")
    
    while True:
        try:
            # 添加更长的随机延迟，模拟人类行为
            time.sleep(random.uniform(3, 6))
            
            response = wcapi.get("orders", params=params)
            
            # 检查响应状态码
            if response.status_code == 403:
                print(f"站点 {site_url} 返回 403 错误，可能启用了安全防护（如 Cloudflare）")
                print(f"响应内容: {response.text[:200]}...")
                
                # 如果是第一次遇到 403 错误，尝试重试
                if retry_count < max_retries:
                    retry_count += 1
                    print(f"第 {retry_count} 次重试...")
                    # 增加延迟时间
                    time.sleep(random.uniform(15, 30))  
                    continue
                else:
                    break
            elif response.status_code == 401:
                print(f"站点 {site_url} 认证失败，请检查 API 密钥是否正确")
                break
            elif response.status_code == 429:  # 速率限制
                print(f"站点 {site_url} 返回速率限制错误，等待后重试...")
                time.sleep(60)  # 等待 1 分钟
                continue
            elif response.status_code == 520:  # Cloudflare 特定错误
                print(f"站点 {site_url} 返回 Cloudflare 520 错误（Web Server Returned an Unknown Error）")
                print(f"这通常表示服务器遇到了问题")
                
                if retry_count < max_retries:
                    retry_count += 1
                    print(f"第 {retry_count} 次重试...")
                    time.sleep(random.uniform(30, 60))  # 更长的延迟
                    continue
                else:
                    break
            elif response.status_code != 200:
                print(f"站点 {site_url} 返回 HTTP {response.status_code} 错误")
                print(f"响应内容: {response.text[:200]}...")
                
                # 如果是服务器错误，尝试重试
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
            retry_count = 0  # 重置重试计数
            
            # 在页面之间添加额外延迟
            time.sleep(random.uniform(2, 5))
        except TimeoutException:
            print(f"获取站点 {site_url} 订单时超时，页码: {page}")
            if retry_count < max_retries:
                retry_count += 1
                print(f"第 {retry_count} 次重试...")
                time.sleep(random.uniform(15, 30))
                continue
            else:
                break
        except HTTPError as e:
            print(f"获取站点 {site_url} 订单时网络错误: {e}")
            if retry_count < max_retries:
                retry_count += 1
                print(f"第 {retry_count} 次重试...")
                time.sleep(random.uniform(15, 30))
                continue
            else:
                break
        except json.JSONDecodeError as e:
            print(f"解析站点 {site_url} 订单数据时出错，可能是返回了非 JSON 内容: {e}")
            print(f"响应内容前200字符: {response.text[:200] if hasattr(response, 'text') else '无响应内容'}")
            if retry_count < max_retries:
                retry_count += 1
                print(f"第 {retry_count} 次重试...")
                time.sleep(random.uniform(15, 30))
                continue
            else:
                break
        except Exception as e:
            print(f"获取站点 {site_url} 订单时发生未知错误: {e}")
            if retry_count < max_retries:
                retry_count += 1
                print(f"第 {retry_count} 次重试...")
                time.sleep(random.uniform(15, 30))
                continue
            else:
                break
    
    return orders

def sync_order_status(wcapi, site_url):
    """同步订单状态"""
    print(f"开始同步站点 {site_url} 的订单状态...")
    
    # 从数据库获取所有订单ID
    order_ids = get_order_ids_from_db(site_url)
    if not order_ids:
        print(f"数据库中没有找到站点 {site_url} 的订单")
        return
    
    print(f"找到 {len(order_ids)} 个订单需要同步状态")
    
    # 分批处理订单状态同步
    batch_size = 10
    for i in range(0, len(order_ids), batch_size):
        batch_ids = order_ids[i:i+batch_size]
        print(f"正在同步第 {i//batch_size + 1} 批订单 ({len(batch_ids)} 个)...")
        
        for order_id in batch_ids:
            try:
                # 添加延迟避免过于频繁的请求
                time.sleep(random.uniform(2, 4))
                
                # 获取单个订单的最新状态
                response = wcapi.get(f"orders/{order_id}")
                
                if response.status_code == 200:
                    order_data = response.json()
                    current_status = order_data.get('status')
                    date_modified = order_data.get('date_modified')
                    
                    # 更新数据库中的订单状态
                    if update_order_status_in_db(order_id, current_status, date_modified):
                        print(f"订单 #{order_id} 状态同步成功: {current_status}")
                    else:
                        print(f"订单 #{order_id} 状态同步失败")
                elif response.status_code == 404:
                    print(f"订单 #{order_id} 在远程站点不存在")
                else:
                    print(f"获取订单 #{order_id} 状态失败: HTTP {response.status_code}")
                    
            except Exception as e:
                print(f"同步订单 #{order_id} 状态时出错: {e}")
                continue
        
        # 批次之间添加更长的延迟
        if i + batch_size < len(order_ids):
            print("等待下一批次...")
            time.sleep(random.uniform(10, 15))

def main(incremental=True, sync_status=False):
    """主函数"""
    print("开始 WooCommerce 订单同步程序...")
    
    # 创建数据库表
    create_orders_table()
    
    for site in sites:
        site_url = site['url']
        print(f"\n处理站点: {site_url}")
        
        # 创建 WooCommerce API 客户端
        wcapi = create_robust_wcapi(site_url, site['ck'], site['cs'], PROXY_CONFIG)
        if not wcapi:
            print(f"无法创建站点 {site_url} 的 API 客户端，跳过")
            continue
        
        try:
            if sync_status:
                # 同步订单状态
                sync_order_status(wcapi, site_url)
            else:
                # 获取订单数据
                last_order_date = None
                if incremental:
                    last_order_date = get_last_order_date_from_db(site_url)
                    if last_order_date:
                        print(f"上次同步的最新订单日期: {last_order_date}")
                
                orders = fetch_orders_incrementally(wcapi, site_url, last_order_date)
                
                if orders:
                    # 为每个订单添加来源站点信息
                    for order in orders:
                        order['source'] = site_url
                    
                    print(f"从站点 {site_url} 获取到 {len(orders)} 个订单")
                    save_orders_to_db(orders)
                else:
                    print(f"站点 {site_url} 没有新订单")
                    
        except Exception as e:
            print(f"处理站点 {site_url} 时发生错误: {e}")
            continue
        
        # 站点之间添加延迟
        print(f"站点 {site_url} 处理完成，等待处理下一个站点...")
        time.sleep(random.uniform(30, 60))
    
    print("\n所有站点处理完成！")

if __name__ == "__main__":
    # 默认运行增量同步
    main(incremental=True, sync_status=False)
    
    # 如果需要同步订单状态，取消注释下面的行
    # main(incremental=True, sync_status=True)
    
    # 如果需要全量同步，取消注释下面的行
    # main(incremental=False, sync_status=False)