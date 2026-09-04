import sqlite3
import unittest
from unittest.mock import Mock, patch

from tracking_format import detect_site_tracking_format


class TrackingFormatDetectionTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self.db.execute(
            """CREATE TABLE orders (
                 source TEXT,
                 status TEXT,
                 meta_data TEXT,
                 line_items TEXT,
                 date_modified TEXT
               )"""
        )

    def tearDown(self):
        self.db.close()

    @patch("requests.get")
    def test_new_site_detects_ast_from_public_namespace(self, get):
        response = Mock(status_code=200)
        response.json.return_value = {
            "namespaces": ["wc/v3", "wc-shipment-tracking/v3"]
        }
        get.return_value = response

        result = detect_site_tracking_format(self.db, "https://vapekram.cz/")

        self.assertEqual("ast", result)
        get.assert_called_once_with(
            "https://vapekram.cz/wp-json/",
            timeout=6,
            headers={"User-Agent": "WooCommerce Order Analysis/1.0"},
        )

    @patch("requests.get")
    def test_new_site_detects_ast_pro_namespace(self, get):
        response = Mock(status_code=200)
        response.json.return_value = {
            "namespaces": ["wc/v3", "wc-ast-pro/v3", "wc-ast/v3"]
        }
        get.return_value = response

        self.assertEqual(
            "ast",
            detect_site_tracking_format(self.db, "https://hifancypolska.pl/"),
        )

    @patch("requests.get")
    def test_local_tracking_history_remains_authoritative(self, get):
        self.db.execute(
            "INSERT INTO orders VALUES (?, ?, ?, ?, ?)",
            (
                "https://history.example",
                "completed",
                "[]",
                '[{"meta_data":[{"key":"_vi_wot_order_item_tracking_data"}]}]',
                "2026-09-04T00:00:00",
            ),
        )

        self.assertEqual(
            "villatheme",
            detect_site_tracking_format(self.db, "https://history.example"),
        )
        get.assert_not_called()

    @patch("requests.get", side_effect=TimeoutError("offline"))
    def test_remote_detection_failure_keeps_safe_unknown_fallback(self, _get):
        self.assertEqual(
            "unknown",
            detect_site_tracking_format(self.db, "https://offline.example"),
        )


if __name__ == "__main__":
    unittest.main()
