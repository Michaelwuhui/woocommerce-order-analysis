#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WooCommerce 订单数据导出到 Excel 文件
从 SQLite 数据库导出订单数据为 Excel 格式
"""

import sqlite3
import pandas as pd
from datetime import datetime
import os

def export_orders_to_excel():
    """将 SQLite 数据库中的订单数据导出为 Excel 文件"""
    
    # 数据库文件路径
    db_file = 'woocommerce_orders.db'
    
    # 检查数据库文件是否存在
    if not os.path.exists(db_file):
        print(f"错误：数据库文件 {db_file} 不存在")
        return
    
    try:
        # 连接到 SQLite 数据库
        print("正在连接到 SQLite 数据库...")
        conn = sqlite3.connect(db_file)
        
        # 查询所有订单数据
        print("正在查询订单数据...")
        query = "SELECT * FROM orders"
        
        # 使用 pandas 读取数据
        df = pd.read_sql_query(query, conn)
        
        # 关闭数据库连接
        conn.close()
        
        print(f"成功读取 {len(df)} 条订单记录")
        
        # 生成带时间戳的文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"woocommerce_orders_{timestamp}.xlsx"
        
        # 导出为 Excel 文件
        print(f"正在导出数据到 Excel 文件: {excel_filename}")
        
        # 使用 ExcelWriter 来更好地控制格式
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            # 写入主要数据
            df.to_excel(writer, sheet_name='订单数据', index=False)
            
            # 获取工作表对象以进行格式化
            worksheet = writer.sheets['订单数据']
            
            # 自动调整列宽
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                # 设置列宽，最大不超过 50
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # 冻结首行（标题行）
            worksheet.freeze_panes = 'A2'
        
        print(f"✅ 成功导出 {len(df)} 条订单数据到 Excel 文件: {excel_filename}")
        
        # 显示一些统计信息
        print("\n📊 数据统计:")
        print(f"   总订单数: {len(df)}")
        
        if 'source' in df.columns:
            print("   按站点分布:")
            source_counts = df['source'].value_counts()
            for source, count in source_counts.items():
                print(f"     {source}: {count} 个订单")
        
        if 'status' in df.columns:
            print("   按状态分布:")
            status_counts = df['status'].value_counts()
            for status, count in status_counts.items():
                print(f"     {status}: {count} 个订单")
        
        if 'total' in df.columns:
            try:
                # 转换为数值类型进行计算
                df['total_numeric'] = pd.to_numeric(df['total'], errors='coerce')
                total_amount = df['total_numeric'].sum()
                print(f"   订单总金额: {total_amount:.2f}")
            except:
                print("   无法计算订单总金额")
        
        print(f"\n📁 文件保存位置: {os.path.abspath(excel_filename)}")
        
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")
    except Exception as e:
        print(f"导出过程中发生错误: {e}")

if __name__ == "__main__":
    export_orders_to_excel()