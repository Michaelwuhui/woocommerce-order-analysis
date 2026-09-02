import sys
sys.path.insert(0, '/www/wwwroot/woo-analysis')
from app import app, get_db_connection
from woocommerce import API

def run_diagnosis():
    with app.app_context():
        conn = get_db_connection()
        # 获取站点信息 (buchmistrz.pl)
        site = conn.execute("SELECT * FROM sites WHERE url LIKE '%buchmistrz.pl%'").fetchone()
        conn.close()
        if not site:
            print("错误: 找不到 buchmistrz.pl 站点")
            return

        order_id = 44951 # 目标订单
        print(f"=== 开始只读诊断订单 #{order_id} (站点: {site['url']}) ===")

        wcapi = API(
            url=site['url'],
            consumer_key=site['consumer_key'],
            consumer_secret=site['consumer_secret'],
            version="wc/v3",
            timeout=30
        )

        print("\n测试读取订单...")
        try:
            resp = wcapi.get(f"orders/{order_id}")
            if resp.status_code == 200:
                print(f"✅ 读取成功! 状态: {resp.json().get('status')}")
            else:
                print(f"❌ 读取失败: {resp.status_code} - {resp.text}")
                return
        except Exception as e:
            print(f"❌ 读取连接错误: {e}")
            return

        print("写权限: 未测试（生产安全只读模式）")

if __name__ == "__main__":
    run_diagnosis()
