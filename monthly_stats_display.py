import sqlite3
import json
import pandas as pd
from datetime import datetime
import argparse
import os

def parse_date(s):
    if not s:
        return None
    try:
        if 'T' in s:
            return datetime.fromisoformat(s.replace('Z', '+00:00'))
        return datetime.strptime(s, '%Y-%m-%d %H:%M:%S')
    except Exception:
        return None

def load_orders(db_path):
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query(
        "SELECT id, date_created, source, status, total, line_items FROM orders",
        conn,
    )
    conn.close()
    return df

def compute_product_quantity(line_items):
    if pd.isna(line_items) or line_items == '':
        return 0
    try:
        items = json.loads(line_items)
        return sum(int(item.get('quantity', 0)) for item in items)
    except Exception:
        return 0

def build_monthly_stats(df):
    df = df.copy()
    df['产品数量'] = df['line_items'].apply(compute_product_quantity)
    df['创建时间'] = df['date_created'].apply(parse_date)
    df = df.dropna(subset=['创建时间'])
    df['月份'] = df['创建时间'].dt.to_period('M').astype(str)
    df['来源网站'] = df['source']
    rows = []
    for (month, source), gdf in df.groupby(['月份', '来源网站']):
        total_orders = len(gdf)
        total_products = int(gdf['产品数量'].sum())
        total_amount = float(pd.to_numeric(gdf['total'], errors='coerce').fillna(0).sum())
        success_mask = ~gdf['status'].isin(['failed', 'cancelled'])
        success_orders = int(success_mask.sum())
        success_products = int(gdf.loc[success_mask, '产品数量'].sum())
        success_amount = float(pd.to_numeric(gdf.loc[success_mask, 'total'], errors='coerce').fillna(0).sum())
        failed_mask = gdf['status'] == 'failed'
        failed_orders = int(failed_mask.sum())
        failed_products = int(gdf.loc[failed_mask, '产品数量'].sum())
        failed_amount = float(pd.to_numeric(gdf.loc[failed_mask, 'total'], errors='coerce').fillna(0).sum())
        cancelled_mask = gdf['status'] == 'cancelled'
        cancelled_orders = int(cancelled_mask.sum())
        cancelled_products = int(gdf.loc[cancelled_mask, '产品数量'].sum())
        cancelled_amount = float(pd.to_numeric(gdf.loc[cancelled_mask, 'total'], errors='coerce').fillna(0).sum())
        completion = round((total_products / 2000) * 100, 2)
        rows.append({
            '月份': month,
            '来源网站': source,
            '总订单数': total_orders,
            '总产品数量': total_products,
            '总订单总额': total_amount,
            '目标完成度(%)': completion,
            '失败订单总数': failed_orders,
            '失败产品数量': failed_products,
            '失败订单总金额': failed_amount,
            '取消订单总数': cancelled_orders,
            '取消产品数量': cancelled_products,
            '取消订单总金额': cancelled_amount,
            '成功订单数': success_orders,
            '成功产品数量': success_products,
            '成功销售金额': success_amount,
        })
    cols = [
        '月份', '来源网站',
        '总订单数', '总产品数量', '总订单总额', '目标完成度(%)',
        '失败订单总数', '失败产品数量', '失败订单总金额',
        '取消订单总数', '取消产品数量', '取消订单总金额',
        '成功订单数', '成功产品数量', '成功销售金额',
    ]
    return pd.DataFrame(rows)[cols]

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--start', type=str)
    parser.add_argument('--db', type=str, default='woocommerce_orders.db')
    args = parser.parse_args()
    df = load_orders(args.db)
    if args.start:
        try:
            dt = datetime.strptime(args.start, '%Y-%m-%d')
            dt_str = dt.strftime('%Y-%m-%d')
            df = df[df['date_created'] >= dt_str]
        except Exception:
            pass
    monthly = build_monthly_stats(df)
    excel_name = 'monthly_stats_from_db.xlsx'
    html_name = 'monthly_stats.html'
    with pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
        monthly.to_excel(writer, sheet_name='月度统计', index=False)
    with open(html_name, 'w', encoding='utf-8') as f:
        f.write(monthly.to_html(index=False))
    print('生成文件:', os.path.abspath(excel_name))
    print('生成文件:', os.path.abspath(html_name))
    print('预览前5行:')
    print(monthly.head().to_string(index=False))

if __name__ == '__main__':
    main()

