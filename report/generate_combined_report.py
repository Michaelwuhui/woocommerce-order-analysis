#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
综合报表生成脚本
整合网站流量数据和订单数据，按站点+月份生成综合分析报表
"""

import pandas as pd
import os
import glob
from datetime import datetime
from collections import defaultdict

# 配置
REPORT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(REPORT_DIR, f'combined_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')

# 站点名称映射（从文件名提取）
SITE_FILE_MAPPING = {
    'strefajednorazowek.pl': ['strefajednorazowek.pl*.csv'],
    'buchmistrz.pl': ['buchmistrz.pl*.csv'],
    'vapepolska.pl': ['vapepolska.pl*.csv'],
    'vapeprime.pl': ['vapeprime.pl.csv'],
    'vapico.pl': ['vapico.pl.csv'],
}

def parse_traffic_csv(filepath, site_name):
    """解析流量CSV文件，返回按月聚合的数据"""
    try:
        # 尝试GBK编码读取
        df = pd.read_csv(filepath, encoding='gbk', skiprows=1)
        
        # 重命名列
        df.columns = ['日期', 'IP数', 'PV', 'UV', '新访客数', '会话数', '跳出率', '平均停留时间']
        
        # 转换日期
        df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
        df = df.dropna(subset=['日期'])
        
        # 创建年月列
        df['年月'] = df['日期'].dt.to_period('M').astype(str)
        
        # 按月聚合
        monthly = df.groupby('年月').agg({
            'UV': 'sum',
            'PV': 'sum',
            '新访客数': 'sum',
            '会话数': 'sum',
            '跳出率': 'mean'
        }).reset_index()
        
        monthly['站点'] = site_name
        return monthly
        
    except Exception as e:
        print(f"读取 {filepath} 失败: {e}")
        return pd.DataFrame()

def load_all_traffic_data():
    """加载所有流量数据"""
    all_traffic = []
    
    for site_name, patterns in SITE_FILE_MAPPING.items():
        for pattern in patterns:
            files = glob.glob(os.path.join(REPORT_DIR, pattern))
            for f in files:
                if 'combined_report' not in f and 'orders_export' not in f:
                    print(f"处理流量文件: {os.path.basename(f)}")
                    traffic_df = parse_traffic_csv(f, site_name)
                    if not traffic_df.empty:
                        all_traffic.append(traffic_df)
    
    if all_traffic:
        return pd.concat(all_traffic, ignore_index=True)
    return pd.DataFrame()

def load_orders_data():
    """加载订单数据"""
    # 查找最新的订单导出文件
    order_files = glob.glob(os.path.join(REPORT_DIR, 'orders_export_*.csv'))
    
    if not order_files:
        print("未找到订单导出文件")
        return pd.DataFrame()
    
    # 使用最新的订单文件
    latest_file = max(order_files, key=os.path.getmtime)
    print(f"使用订单文件: {os.path.basename(latest_file)}")
    
    try:
        df = pd.read_csv(latest_file, encoding='utf-8-sig')
        
        # 转换日期
        df['创建日期'] = pd.to_datetime(df['创建日期'], errors='coerce')
        df = df.dropna(subset=['创建日期'])
        
        # 创建年月列
        df['年月'] = df['创建日期'].dt.to_period('M').astype(str)
        
        # 排除无效订单状态
        invalid_statuses = ['已取消', '失败', 'checkout-draft', 'trash', 'cheat']
        df = df[~df['状态'].isin(invalid_statuses)]
        
        # 按站点和月份聚合
        monthly = df.groupby(['来源', '年月']).agg({
            '订单号': 'count',
            '订单金额': 'sum',
            '净额': 'sum',
            '¥净额': 'sum',
            '总数量': 'sum'
        }).reset_index()
        
        monthly.columns = ['站点', '年月', '订单数', '订单金额', '净额', '净额CNY', '产品数量']
        
        return monthly
        
    except Exception as e:
        print(f"读取订单数据失败: {e}")
        return pd.DataFrame()

def merge_and_calculate(traffic_df, orders_df):
    """合并数据并计算指标"""
    
    if traffic_df.empty and orders_df.empty:
        print("没有可用数据")
        return pd.DataFrame()
    
    # 按站点+年月聚合流量数据（可能有多个文件覆盖同一时期）
    if not traffic_df.empty:
        traffic_agg = traffic_df.groupby(['站点', '年月']).agg({
            'UV': 'sum',
            'PV': 'sum',
            '新访客数': 'sum',
            '会话数': 'sum',
            '跳出率': 'mean'
        }).reset_index()
    else:
        traffic_agg = pd.DataFrame()
    
    # 合并数据
    if not traffic_agg.empty and not orders_df.empty:
        merged = pd.merge(
            traffic_agg,
            orders_df,
            on=['站点', '年月'],
            how='outer'
        )
    elif not traffic_agg.empty:
        merged = traffic_agg
    else:
        merged = orders_df
    
    # 填充缺失值
    numeric_cols = ['UV', 'PV', '新访客数', '会话数', '订单数', '订单金额', '净额', '净额CNY', '产品数量']
    for col in numeric_cols:
        if col in merged.columns:
            merged[col] = merged[col].fillna(0)
    
    # 计算转化率和客单价
    if 'UV' in merged.columns and '订单数' in merged.columns:
        merged['转化率'] = merged.apply(
            lambda x: f"{(x['订单数'] / x['UV'] * 100):.2f}%" if x['UV'] > 0 else "N/A",
            axis=1
        )
    
    if '净额CNY' in merged.columns and '订单数' in merged.columns:
        merged['客单价CNY'] = merged.apply(
            lambda x: round(x['净额CNY'] / x['订单数'], 2) if x['订单数'] > 0 else 0,
            axis=1
        )
    
    # 排序
    merged = merged.sort_values(['站点', '年月'])
    
    return merged

def generate_pivot_report(merged_df):
    """生成透视表格式的报表（按用户图片中的布局）"""
    
    if merged_df.empty:
        return pd.DataFrame()
    
    # 获取所有站点和月份
    sites = merged_df['站点'].unique()
    months = sorted(merged_df['年月'].unique())
    
    # 指标列表
    metrics = ['UV', 'PV', '订单数', '订单金额', '净额CNY', '转化率', '客单价CNY']
    
    # 构建透视表数据
    rows = []
    for site in sites:
        site_data = merged_df[merged_df['站点'] == site]
        
        for metric in metrics:
            row = {'站点': site, '指标': metric}
            
            for month in months:
                month_data = site_data[site_data['年月'] == month]
                if not month_data.empty and metric in month_data.columns:
                    value = month_data[metric].values[0]
                    if metric in ['订单金额', '净额CNY', '客单价CNY']:
                        row[month] = f"{value:,.2f}" if isinstance(value, (int, float)) else value
                    elif metric in ['UV', 'PV', '订单数']:
                        row[month] = int(value) if isinstance(value, (int, float)) else value
                    else:
                        row[month] = value
                else:
                    row[month] = ''
            
            # 计算合计/平均
            if metric in ['UV', 'PV', '订单数', '订单金额', '净额CNY']:
                if metric in site_data.columns:
                    row['合计'] = f"{site_data[metric].sum():,.2f}" if metric in ['订单金额', '净额CNY'] else int(site_data[metric].sum())
            elif metric == '客单价CNY':
                total_revenue = site_data['净额CNY'].sum() if '净额CNY' in site_data.columns else 0
                total_orders = site_data['订单数'].sum() if '订单数' in site_data.columns else 0
                row['平均'] = f"{total_revenue / total_orders:,.2f}" if total_orders > 0 else 'N/A'
            elif metric == '转化率':
                total_uv = site_data['UV'].sum() if 'UV' in site_data.columns else 0
                total_orders = site_data['订单数'].sum() if '订单数' in site_data.columns else 0
                row['平均'] = f"{(total_orders / total_uv * 100):.2f}%" if total_uv > 0 else 'N/A'
            
            rows.append(row)
    
    return pd.DataFrame(rows)

def main():
    print("=" * 60)
    print("综合报表生成")
    print("=" * 60)
    
    # 加载数据
    print("\n1. 加载流量数据...")
    traffic_df = load_all_traffic_data()
    print(f"   流量数据行数: {len(traffic_df)}")
    
    print("\n2. 加载订单数据...")
    orders_df = load_orders_data()
    print(f"   订单聚合行数: {len(orders_df)}")
    
    # 合并和计算
    print("\n3. 合并数据并计算指标...")
    merged_df = merge_and_calculate(traffic_df, orders_df)
    
    if merged_df.empty:
        print("没有生成数据")
        return
    
    # 生成透视表报表
    print("\n4. 生成透视表报表...")
    pivot_df = generate_pivot_report(merged_df)
    
    # 保存报表
    output_file = os.path.join(REPORT_DIR, f'combined_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
    pivot_df.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"\n✅ 报表已保存: {output_file}")
    
    # 同时保存一份详细数据
    detail_file = os.path.join(REPORT_DIR, f'combined_detail_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
    merged_df.to_csv(detail_file, index=False, encoding='utf-8-sig')
    print(f"✅ 详细数据已保存: {detail_file}")
    
    # 打印预览
    print("\n" + "=" * 60)
    print("报表预览（透视表格式）:")
    print("=" * 60)
    print(pivot_df.to_string(index=False))

if __name__ == '__main__':
    main()
