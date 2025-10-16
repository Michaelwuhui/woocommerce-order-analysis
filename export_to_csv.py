import mysql.connector
import csv
import os
from datetime import datetime

def export_orders_to_csv():
    """
    将MySQL数据库中的订单数据导出为CSV文件
    """
    # 数据库连接配置
    db_config = {
        'host': 'localhost',
        'user': 'root',
        'password': '1207wlp@',
        'database': 'woocommerce_orders',
        'charset': 'utf8mb4'
    }
    
    try:
        # 连接数据库
        print("正在连接数据库...")
        connection = mysql.connector.connect(**db_config)
        cursor = connection.cursor()
        
        # 查询所有订单数据
        query = "SELECT * FROM orders"
        cursor.execute(query)
        
        # 获取列名
        column_names = [desc[0] for desc in cursor.description]
        
        # 获取所有数据
        rows = cursor.fetchall()
        
        # 生成CSV文件名（包含时间戳）
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_filename = f"woocommerce_orders_{timestamp}.csv"
        csv_filepath = os.path.join(os.getcwd(), csv_filename)
        
        # 写入CSV文件
        print(f"正在导出数据到 {csv_filename}...")
        with open(csv_filepath, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            
            # 写入列标题
            writer.writerow(column_names)
            
            # 写入数据行
            for row in rows:
                # 处理可能的None值
                processed_row = [str(cell) if cell is not None else '' for cell in row]
                writer.writerow(processed_row)
        
        print(f"导出完成！")
        print(f"文件路径: {csv_filepath}")
        print(f"总共导出 {len(rows)} 条订单记录")
        print(f"包含字段: {', '.join(column_names)}")
        
    except mysql.connector.Error as err:
        print(f"数据库错误: {err}")
    except Exception as e:
        print(f"导出过程中发生错误: {e}")
    finally:
        # 关闭数据库连接
        if 'connection' in locals() and connection.is_connected():
            cursor.close()
            connection.close()
            print("数据库连接已关闭")

if __name__ == "__main__":
    export_orders_to_csv()