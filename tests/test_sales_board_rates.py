import os
import sqlite3
import sys
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from sales_board_rates import (  # noqa: E402
    load_monthly_receipt_rates,
    resolve_sales_board_rate,
)


class SalesBoardReceiptRateTests(unittest.TestCase):
    def setUp(self):
        self.conn = sqlite3.connect(":memory:")
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(
            """
            CREATE TABLE partners (
                id INTEGER PRIMARY KEY,
                name TEXT,
                currency TEXT
            );
            CREATE TABLE partner_receipts (
                id INTEGER PRIMARY KEY,
                partner_id INTEGER,
                receipt_date TEXT,
                amount_pln REAL,
                exchange_rate_cny REAL,
                amount_cny REAL
            );
            INSERT INTO partners VALUES (1, '波兰合伙人', 'PLN');
            INSERT INTO partners VALUES (2, '澳洲合伙人', 'AUD');
            INSERT INTO partner_receipts VALUES
                (1, 1, '2026-06-05', 8500, 1.82, 15470),
                (2, 1, '2026-06-20', 33000, 1.78, 58740),
                (3, 2, '2026-06-06', 1792, 4.72, 8458.24),
                (4, 2, '2026-06-21', 3111, 4.71, 14652.81),
                (5, 1, '2026-05-30', 999, 9.99, 9980.01);
            """
        )

    def tearDown(self):
        self.conn.close()

    def test_monthly_rate_is_weighted_by_native_receipt_amount(self):
        rates = load_monthly_receipt_rates(self.conn, "2026-06")

        self.assertAlmostEqual(rates["PLN"]["rate"], 74210 / 41500)
        self.assertAlmostEqual(rates["AUD"]["rate"], 23111.05 / 4903)
        self.assertEqual(rates["PLN"]["receipt_count"], 2)
        self.assertEqual(rates["PLN"]["partner_names"], ["波兰合伙人"])

    def test_receipt_rate_wins_then_custom_then_system(self):
        receipts = load_monthly_receipt_rates(self.conn, "2026-06")

        rate, source = resolve_sales_board_rate(
            "PLN", receipts, {"PLN": 1.75}, 1.82288
        )
        self.assertAlmostEqual(rate, 74210 / 41500)
        self.assertEqual(source, "receipt")

        rate, source = resolve_sales_board_rate(
            "EUR", receipts, {"EUR": 8.1}, 8.2
        )
        self.assertEqual((rate, source), (8.1, "override"))

        rate, source = resolve_sales_board_rate(
            "USD", receipts, {}, 7.2
        )
        self.assertEqual((rate, source), (7.2, "system"))


if __name__ == "__main__":
    unittest.main()
