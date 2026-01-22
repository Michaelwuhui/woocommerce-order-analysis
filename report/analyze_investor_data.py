import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import glob
import os

# Set style
sns.set(style="whitegrid")
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial'] # Try to support Chinese if available, fallback to Arial
plt.rcParams['axes.unicode_minus'] = False

def parse_traffic_data():
    all_files = glob.glob('*.csv')
    traffic_files = [f for f in all_files if 'orders_export' not in f and 'investor_data' not in f]
    
    df_list = []
    
    for filename in traffic_files:
        try:
            site_name = filename.split('.pl')[0] + '.pl'
            
            # Try reading with different encodings
            try:
                df = pd.read_csv(filename, header=1, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(filename, header=1, encoding='gbk')
            
            # Rename columns from Chinese to English standard
            # 时间,IP数,访问数(PV),访客数(UV),新访客数,会话数,跳出率,平均访问时长
            column_map = {
                '时间': 'Date',
                '会话数': 'Sessions',
                '访客数(UV)': 'Users'
            }
            df.rename(columns=column_map, inplace=True)
            
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'])
                 
            df['Site'] = site_name
            df_list.append(df)
        except Exception as e:
            print(f"Error reading {filename}: {e}")
            
    if not df_list:
        return pd.DataFrame()
        
    return pd.concat(df_list, ignore_index=True)

def parse_order_data(order_file):
    # Read with header (default)
    df = pd.read_csv(order_file)
    
    # Rename valuable columns
    # 订单号,创建日期,状态,来源,负责人,客户姓名,客户邮箱,产品明细,总数量,运费,订单金额,净额,汇率,货币,¥净额
    rename_map = {
        '创建日期': 'Date',
        '来源': 'Site',
        '订单号': 'Order Number',
        '¥净额': 'Order Total CNY',
        '净额': 'Order Net',
        '状态': 'Status'
    }
    # It allows for partial renaming
    df.rename(columns=rename_map, inplace=True)
    
    # Parse Date
    if 'Date' in df.columns:
        # 2026-01-21T21:03:39 format
        df['Date'] = pd.to_datetime(df['Date']).dt.normalize()
    
    return df

def main():
    print("Starting analysis...")
    
    # 1. Load Traffic
    traffic_df = parse_traffic_data()
    print(f"Loaded {len(traffic_df)} traffic rows.")
    if not traffic_df.empty:
        # Normalize Date
        traffic_df['Date'] = pd.to_datetime(traffic_df['Date'])
        # Aggregate strings like '1,234' to floats if needed
        for col in ['Users', 'Sessions', 'New users']: # Typical GA cols
            if col in traffic_df.columns and traffic_df[col].dtype == object:
                 traffic_df[col] = traffic_df[col].str.replace(',', '').astype(float)
    
    # 2. Load Orders
    order_file = glob.glob('orders_export_*.csv')[0]
    orders_df = parse_order_data(order_file)
    print(f"Loaded {len(orders_df)} orders.")

    # 3. Aggregate Orders by Date + Site
    daily_orders = orders_df.groupby(['Date', 'Site']).size().reset_index(name='Order Count')
    daily_revenue = orders_df.groupby(['Date', 'Site'])['Order Total CNY'].sum().reset_index(name='Revenue CNY')
    
    orders_agg = pd.merge(daily_orders, daily_revenue, on=['Date', 'Site'])
    
    # 4. Merge
    # Note: Traffic data might determine the date range foundation
    if not traffic_df.empty:
        merged_df = pd.merge(traffic_df, orders_agg, on=['Date', 'Site'], how='outer').fillna(0)
    else:
        merged_df = orders_agg
        
    merged_df.sort_values('Date', inplace=True)
    
    # 5. Calculate Metrics
    if 'Sessions' in merged_df.columns:
        merged_df['Conversion Rate %'] = (merged_df['Order Count'] / merged_df['Sessions'] * 100).fillna(0)
    
    # 6. Export Summary
    output_file = 'investor_report_data.xlsx'
    merged_df.to_excel(output_file, index=False)
    print(f"Saved summary to {output_file}")
    
    # 7. Visualization
    # Group by Date (All Sites) for high level trend
    if 'Sessions' in merged_df.columns:
        total_trend = merged_df.groupby('Date')[['Order Count', 'Revenue CNY', 'Sessions']].sum().reset_index()
    else:
        total_trend = merged_df.groupby('Date')[['Order Count', 'Revenue CNY']].sum().reset_index()
    
    fig, ax1 = plt.subplots(figsize=(12, 6))
    
    color = 'tab:blue'
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Traffic (Sessions)', color=color)
    if 'Sessions' in total_trend.columns:
        ax1.plot(total_trend['Date'], total_trend['Sessions'], color=color, label='Sessions')
    ax1.tick_params(axis='y', labelcolor=color)
    
    ax2 = ax1.twinx()  # instantiate a second axes that shares the same x-axis
    
    color = 'tab:red'
    ax2.set_ylabel('Orders', color=color)  # we already handled the x-label with ax1
    ax2.plot(total_trend['Date'], total_trend['Order Count'], color=color, linestyle='--', label='Orders')
    ax2.tick_params(axis='y', labelcolor=color)
    
    plt.title('Traffic vs Orders Trend (All Sites)')
    fig.tight_layout()  # otherwise the right y-label is slightly clipped
    plt.savefig('traffic_vs_orders_trend.png')
    print("Saved traffic_vs_orders_trend.png")
    
    # Conversion Rate
    if 'Conversion Rate %' in merged_df.columns:
        plt.figure(figsize=(12, 6))
        # Pivot for per-site line chart? Or just total?
        # Let's do Total first
        total_cr = total_trend['Order Count'] / total_trend['Sessions'] * 100
        plt.plot(total_trend['Date'], total_cr, marker='o', linestyle='-')
        plt.title('Daily Conversion Rate %')
        plt.ylabel('CR %')
        plt.savefig('conversion_rate_trend.png')
        print("Saved conversion_rate_trend.png")

if __name__ == "__main__":
    main()
