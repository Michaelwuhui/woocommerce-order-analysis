"""Read-only WMS credential/connectivity check. Never prints secret or stock rows."""

import os
import hashlib
import sqlite3
import sys
import uuid

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from hungary_wms import WmsClient
from fulfillment_common import json_dump, redact


result = WmsClient().inventory("HU01")
if "--audit" in sys.argv:
    request_data = {"storehouseCode": "HU01"}
    conn = sqlite3.connect("woocommerce_orders.db")
    try:
        conn.execute(
            '''INSERT INTO oms_external_api_calls
               (correlation_id, provider, operation, method, endpoint,
                request_hash, request_redacted, response_http_code,
                response_code, response_redacted, duration_ms, attempt, outcome)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (
                f"validation-{uuid.uuid4().hex}", "hungary_wms", "inventory_validation",
                result.method, result.endpoint,
                hashlib.sha256(json_dump(request_data).encode()).hexdigest(),
                json_dump(request_data), result.http_status, str(result.business_code),
                json_dump(redact(result.raw)), result.duration_ms, 1, "success",
            ),
        )
        conn.commit()
    finally:
        conn.close()
if isinstance(result.data, list):
    shape = "list"
    count = len(result.data)
elif isinstance(result.data, dict):
    shape = "object"
    count = len(result.data)
else:
    shape = type(result.data).__name__
    count = 0

print(
    f"wms_readonly=ok http={result.http_status} business_code={result.business_code} "
    f"warehouse=HU01 data_shape={shape} top_level_count={count}"
)
