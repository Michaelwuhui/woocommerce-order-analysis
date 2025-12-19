from app import app, get_customer_details
import sys

# Mock current_user for login_required
from flask_login import current_user
from unittest.mock import MagicMock

# Create a request context
with app.test_request_context():
    try:
        # Call the function directly
        # We need to mock login_required or bypass it? 
        # Since we import the function, the decorator wraps it.
        # But we can't easily bypass the decorator on an imported function.
        # Instead, let's just use the test client.
        
        client = app.test_client()
        # We need to be logged in.
        # Let's try to mock the login.
        with client.session_transaction() as sess:
            sess['_user_id'] = '1'
            
        response = client.get('/api/customer/wicio610%40gmail.com')
        print(f"Status Code: {response.status_code}")
        if response.status_code != 200:
            print("Response Data:")
            print(response.data.decode('utf-8'))
        else:
            print("Success!")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
