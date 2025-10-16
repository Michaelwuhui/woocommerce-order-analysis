import pandas as pd

# 读取生成的Excel文件
filename = "improved_sales_analysis_20251015_122326.xlsx"

print("📊 检查Excel文件内容...")
print("="*60)

# 读取销售分析报告工作表
df_main = pd.read_excel(filename, sheet_name='销售分析报告')
print("🔍 销售分析报告工作表:")
print(f"   行数: {len(df_main)}")
print(f"   列数: {len(df_main.columns)}")
print(f"   列名: {list(df_main.columns)}")

print("\n📋 前5行数据:")
print(df_main.head())

# 读取网站汇总工作表
df_summary = pd.read_excel(filename, sheet_name='网站汇总')
print("\n🔍 网站汇总工作表:")
print(f"   行数: {len(df_summary)}")
print(f"   列数: {len(df_summary.columns)}")

print("\n📋 汇总数据:")
print(df_summary)

print("\n✅ Excel文件检查完成!")