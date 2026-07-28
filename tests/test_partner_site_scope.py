import os
import sqlite3
import sys
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from partner_site_scope import (  # noqa: E402
    get_partner_site_scope,
    init_partner_site_scope,
    replace_partner_site_scope,
)


class PartnerSiteScopeTests(unittest.TestCase):
    def setUp(self):
        self.conn = sqlite3.connect(":memory:")
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(
            """
            CREATE TABLE partners (
                id INTEGER PRIMARY KEY,
                name TEXT
            );
            CREATE TABLE sites (
                id INTEGER PRIMARY KEY,
                url TEXT,
                country TEXT
            );
            CREATE TABLE partner_sites (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                partner_id INTEGER,
                site_id INTEGER,
                UNIQUE(partner_id, site_id)
            );
            INSERT INTO partners VALUES (1, '波兰合伙人');
            INSERT INTO partners VALUES (2, '其他合伙人');
            INSERT INTO sites VALUES (1, 'https://a.pl', 'PL');
            INSERT INTO sites VALUES (2, 'https://b.pl', 'PL');
            INSERT INTO sites VALUES (3, 'https://c.pl', 'PL');
            INSERT INTO sites VALUES (4, 'https://a.au', 'AU');
            INSERT INTO partner_sites (partner_id, site_id) VALUES (1, 1);
            """
        )
        init_partner_site_scope(self.conn)

    def tearDown(self):
        self.conn.close()

    def test_country_grant_inherits_future_sites_and_allows_exclusion(self):
        result = replace_partner_site_scope(
            self.conn,
            1,
            country_grants=["pl"],
            site_exclusions=[2],
        )
        self.assertEqual(result["granted_countries"], ["PL"])
        self.assertEqual(result["effective_site_ids"], [1, 3])

        self.conn.execute(
            "INSERT INTO sites VALUES (5, 'https://future.pl', 'PL')"
        )
        scope = get_partner_site_scope(self.conn, 1)
        self.assertEqual(scope["effective_site_ids"], {1, 3, 5})
        self.assertEqual(scope["excluded_site_ids"], {2})

    def test_explicit_sites_outside_country_grant_are_preserved(self):
        result = replace_partner_site_scope(
            self.conn,
            1,
            site_ids=[4],
            country_grants=["PL"],
            site_exclusions=[2],
        )
        self.assertEqual(result["explicit_site_ids"], [4])
        self.assertEqual(result["effective_site_ids"], [1, 3, 4])

    def test_effective_site_cannot_belong_to_two_partners(self):
        replace_partner_site_scope(
            self.conn,
            1,
            country_grants=["PL"],
        )
        with self.assertRaisesRegex(ValueError, "不能重复计入"):
            replace_partner_site_scope(
                self.conn,
                2,
                site_ids=[1],
            )
        self.assertEqual(
            get_partner_site_scope(self.conn, 2)["effective_site_ids"],
            set(),
        )


if __name__ == "__main__":
    unittest.main()
