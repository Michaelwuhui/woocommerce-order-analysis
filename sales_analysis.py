import sqlite3
import pandas as pd
from datetime import datetime
import json

def analyze_woocommerce_data():
    """分析WooCommerce订单数据并生成汇总报告"""
    
    # 连接到SQLite数据库
    conn = sqlite3.connect('woocommerce_orders.db')
    
    try:
        # 查询所有订单数据
        query = """
        SELECT 
            id,
            source,
            status,
            total,
            line_items,
            date_created,
            date_completed,
            shipping_lines
        FROM orders
        """
        
        df = pd.read_sql_query(query, conn)
        
        # 解析line_items来计算产品数量
        def count_products(line_items_json):
            try:
                if line_items_json:
                    items = json.loads(line_items_json)
                    return sum(int(item.get('quantity', 0)) for item in items)
                return 0
            except:
                return 0
        
        df['product_count'] = df['line_items'].apply(count_products)
        
        # 按站点分组分析
        sites = ['strefajednorazowek.pl', 'buchmistrz.pl']
        
        analysis_results = []
        
        for site in sites:
            site_data = df[df['source'].str.contains(site, na=False)]
            
            if len(site_data) == 0:
                continue
            
            # 计算各项指标
            total_sales = site_data['total'].astype(float).sum()
            total_orders = len(site_data)
            total_products = site_data['product_count'].sum()
            
            # 成功签收 (completed状态)
            completed_orders = site_data[site_data['status'] == 'completed']
            successful_delivery_count = len(completed_orders)
            successful_delivery_sales = completed_orders['total'].astype(float).sum()
            
            # 发货未签收 (processing, on-hold状态)
            shipped_not_delivered = site_data[site_data['status'].isin(['processing', 'on-hold'])]
            shipped_not_delivered_count = len(shipped_not_delivered)
            shipped_not_delivered_sales = shipped_not_delivered['total'].astype(float).sum()
            
            # 缺货/取消 (cancelled, failed状态)
            out_of_stock = site_data[site_data['status'].isin(['cancelled', 'failed'])]
            out_of_stock_count = len(out_of_stock)
            out_of_stock_sales = out_of_stock['total'].astype(float).sum()
            
            analysis_results.append({
                'site': site,
                'total_sales': total_sales,
                'total_orders': total_orders,
                'total_products': total_products,
                'successful_delivery_sales': successful_delivery_sales,
                'successful_delivery_count': successful_delivery_count,
                'shipped_not_delivered_sales': shipped_not_delivered_sales,
                'shipped_not_delivered_count': shipped_not_delivered_count,
                'out_of_stock_sales': out_of_stock_sales,
                'out_of_stock_count': out_of_stock_count
            })
        
        # 计算汇总
        total_summary = {
            'site': '汇总',
            'total_sales': sum(r['total_sales'] for r in analysis_results),
            'total_orders': sum(r['total_orders'] for r in analysis_results),
            'total_products': sum(r['total_products'] for r in analysis_results),
            'successful_delivery_sales': sum(r['successful_delivery_sales'] for r in analysis_results),
            'successful_delivery_count': sum(r['successful_delivery_count'] for r in analysis_results),
            'shipped_not_delivered_sales': sum(r['shipped_not_delivered_sales'] for r in analysis_results),
            'shipped_not_delivered_count': sum(r['shipped_not_delivered_count'] for r in analysis_results),
            'out_of_stock_sales': sum(r['out_of_stock_sales'] for r in analysis_results),
            'out_of_stock_count': sum(r['out_of_stock_count'] for r in analysis_results)
        }
        
        analysis_results.append(total_summary)
        
        return analysis_results
        
    finally:
        conn.close()

def create_excel_report(analysis_results):
    """创建Excel报告，按照用户提供的格式"""
    
    # 创建DataFrame
    data = []
    
    for result in analysis_results:
        data.append([
            result['site'],
            f"{result['total_sales']:.2f}",
            result['total_orders'],
            result['total_products'],
            f"{result['successful_delivery_sales']:.2f}",
            result['successful_delivery_count'],
            f"{result['shipped_not_delivered_sales']:.2f}",
            result['shipped_not_delivered_count'],
            f"{result['out_of_stock_sales']:.2f}",
            result['out_of_stock_count']
        ])
    
    # 创建DataFrame，列名与用户模板一致
    columns = [
        'A',  # 站点名称
        'B',  # 总销售
        'C',  # 订单数量
        'D',  # 产品数量
        'E',  # 成功签收
        'F',  # 订单数量
        'G',  # 发货未签收
        'H',  # 订单数量
        'I',  # 缺货
        'J'   # 订单数量
    ]
    
    df = pd.DataFrame(data, columns=columns)
    
    # 生成文件名
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'sales_analysis_report_{timestamp}.xlsx'
    
    # 创建Excel文件
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # 创建标题行
        header_data = [
            ['', '总销售', '订单数量', '产品数量', '成功签收', '订单数量', '发货未签收', '订单数量', '缺货', '订单数量']
        ]
        
        header_df = pd.DataFrame(header_data, columns=columns)
        
        # 写入标题
        header_df.to_excel(writer, sheet_name='销售数据分析', index=False, header=False, startrow=1)
        
        # 写入数据
        df.to_excel(writer, sheet_name='销售数据分析', index=False, header=False, startrow=2)
        
        # 获取工作表进行格式化
        worksheet = writer.sheets['销售数据分析']
        
        # 添加主标题
        worksheet['E1'] = '9月份销售数据'
        
        # 设置列宽
        column_widths = [25, 12, 12, 12, 12, 12, 12, 12, 12, 12]
        for i, width in enumerate(column_widths):
            worksheet.column_dimensions[chr(65 + i)].width = width
        
        # 冻结首行
        worksheet.freeze_panes = 'A3'
    
    return filename

def main():
    """主函数"""
    print("开始分析WooCommerce订单数据...")
    
    # 分析数据
    results = analyze_woocommerce_data()
    
    # 打印分析结果
    print("\n=== 销售数据分析结果 ===")
    print(f"{'站点':<25} {'总销售':<12} {'订单数':<8} {'产品数':<8} {'成功签收':<12} {'订单数':<8} {'发货未签收':<12} {'订单数':<8} {'缺货':<12} {'订单数':<8}")
    print("-" * 140)
    
    for result in results:
        print(f"{result['site']:<25} "
              f"{result['total_sales']:<12.2f} "
              f"{result['total_orders']:<8} "
              f"{result['total_products']:<8} "
              f"{result['successful_delivery_sales']:<12.2f} "
              f"{result['successful_delivery_count']:<8} "
              f"{result['shipped_not_delivered_sales']:<12.2f} "
              f"{result['shipped_not_delivered_count']:<8} "
              f"{result['out_of_stock_sales']:<12.2f} "
              f"{result['out_of_stock_count']:<8}")
    
    # 创建Excel报告
    filename = create_excel_report(results)
    print(f"\nExcel报告已生成: {filename}")
    
    return results

if __name__ == "__main__":
    main()