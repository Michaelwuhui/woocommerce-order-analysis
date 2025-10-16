import sqlite3
import json
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

def analyze_line_items():
    """分析line_items字段，计算产品数量统计"""
    
    # 连接数据库
    conn = sqlite3.connect('woocommerce_orders.db')
    cursor = conn.cursor()
    
    print("正在分析line_items字段...")
    
    # 查看表结构
    cursor.execute('PRAGMA table_info(orders)')
    columns = cursor.fetchall()
    print('\n数据库表结构:')
    for col in columns:
        print(f'  {col[1]} ({col[2]})')
    
    print('\n' + '='*50)
    
    # 获取所有订单的line_items数据
    cursor.execute('SELECT id, line_items, date_created, source FROM orders WHERE line_items IS NOT NULL')
    rows = cursor.fetchall()
    
    print(f'找到 {len(rows)} 个包含line_items的订单')
    
    # 分析数据
    analysis_data = []
    total_products_sold = 0
    
    for order_id, line_items, date_created, source in rows:
        try:
            items = json.loads(line_items)
            
            # 计算该订单的统计数据
            product_count = len(items)  # 产品种类数
            total_quantity = sum(item.get('quantity', 0) for item in items)  # 总产品数量
            total_products_sold += total_quantity
            
            # 提取产品信息
            product_names = []
            product_skus = []
            for item in items:
                product_names.append(item.get('name', '未知产品'))
                product_skus.append(item.get('sku', '无SKU'))
            
            analysis_data.append({
                'order_id': order_id,
                'source': source,
                'date_created': date_created,
                'product_types_count': product_count,  # 产品种类数
                'total_quantity': total_quantity,      # 产品数量
                'product_names': ' | '.join(product_names),
                'product_skus': ' | '.join(product_skus),
                'line_items_raw': line_items[:200] + '...' if len(line_items) > 200 else line_items
            })
            
        except json.JSONDecodeError as e:
            print(f'订单 {order_id} JSON解析错误: {e}')
            analysis_data.append({
                'order_id': order_id,
                'source': source,
                'date_created': date_created,
                'product_types_count': 0,
                'total_quantity': 0,
                'product_names': 'JSON解析错误',
                'product_skus': 'JSON解析错误',
                'line_items_raw': line_items[:200] + '...' if len(line_items) > 200 else line_items
            })
    
    conn.close()
    
    # 显示前几个示例
    print('\nline_items字段分析示例:')
    for i, data in enumerate(analysis_data[:5], 1):
        print(f'\n订单 {data["order_id"]}:')
        print(f'  来源: {data["source"]}')
        print(f'  日期: {data["date_created"]}')
        print(f'  产品种类数: {data["product_types_count"]}')
        print(f'  产品总数量: {data["total_quantity"]}')
        print(f'  产品名称: {data["product_names"][:100]}...' if len(data["product_names"]) > 100 else f'  产品名称: {data["product_names"]}')
    
    # 统计汇总
    print(f'\n' + '='*50)
    print('统计汇总:')
    print(f'总订单数: {len(analysis_data)}')
    print(f'总产品销售数量: {total_products_sold} 支')
    
    # 按网站统计
    site_stats = {}
    for data in analysis_data:
        site = data['source']
        if site not in site_stats:
            site_stats[site] = {'orders': 0, 'quantity': 0}
        site_stats[site]['orders'] += 1
        site_stats[site]['quantity'] += data['total_quantity']
    
    print('\n按网站统计:')
    for site, stats in site_stats.items():
        print(f'  {site}: {stats["orders"]} 订单, {stats["quantity"]} 支产品')
    
    # 月度统计
    monthly_stats = {}
    for data in analysis_data:
        try:
            # 解析日期
            date_str = data['date_created']
            if 'T' in date_str:
                date_obj = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            else:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
            
            month_key = date_obj.strftime('%Y-%m')
            if month_key not in monthly_stats:
                monthly_stats[month_key] = {'orders': 0, 'quantity': 0}
            monthly_stats[month_key]['orders'] += 1
            monthly_stats[month_key]['quantity'] += data['total_quantity']
        except Exception as e:
            print(f'日期解析错误: {data["date_created"]} - {e}')
    
    print('\n月度销量统计:')
    for month, stats in sorted(monthly_stats.items()):
        progress = (stats['quantity'] / 2000) * 100
        print(f'  {month}: {stats["quantity"]} 支产品 ({stats["orders"]} 订单) - 目标完成度: {progress:.1f}%')
    
    return analysis_data, total_products_sold, site_stats, monthly_stats

if __name__ == "__main__":
    analysis_data, total_products_sold, site_stats, monthly_stats = analyze_line_items()