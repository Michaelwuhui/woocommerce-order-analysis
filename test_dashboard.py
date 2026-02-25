import sys
sys.path.append('/www/wwwroot/woo-analysis')
from app import app, get_db_connection, User
from flask_login import login_user

def test_dashboard():
    with app.test_request_context('/'):
        # Mock session and login
        app.preprocess_request()
        conn = get_db_connection()
        user_row = conn.execute("SELECT * FROM users WHERE role='admin' LIMIT 1").fetchone()
        if not user_row:
            user_row = conn.execute("SELECT * FROM users LIMIT 1").fetchone()
        conn.close()
        
        user = User(user_row['id'], user_row['username'], user_row['name'], user_row['role'])
        login_user(user)
        
        try:
            from app import dashboard
            resp = dashboard()
            print("Success!")
        except Exception as e:
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    with app.test_client() as c:
        test_dashboard()
