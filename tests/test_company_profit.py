import os
import sqlite3
import sys
import tempfile
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from company_profit import (  # noqa: E402
    _revenue_ladder,
    build_company_profit_summary,
    init_company_profit_tables,
)


class CompanyProfitTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.tempdir.name, "profit-test.db")
        conn = self.get_conn()
        conn.executescript(
            """
            CREATE TABLE users (
                id INTEGER PRIMARY KEY,
                username TEXT,
                name TEXT,
                role TEXT
            );
            INSERT INTO users VALUES (1, 'admin', '管理员', 'admin');
            INSERT INTO users VALUES (2, 'michael', '吴辉', 'viewer');
            INSERT INTO users VALUES (3, 'fiona', '付肖肖', 'admin');

            CREATE TABLE sites (
                id INTEGER PRIMARY KEY,
                url TEXT,
                country TEXT,
                manager TEXT,
                cod_on_hold_is_shipped INTEGER DEFAULT 0
            );
            INSERT INTO sites VALUES (1, 'https://pl.example', 'PL', '吴辉', 1);
            INSERT INTO sites VALUES (2, 'https://au.example', 'AU', '付肖肖', 0);

            CREATE TABLE partners (
                id INTEGER PRIMARY KEY,
                name TEXT,
                currency TEXT,
                cost_ratio REAL,
                partner_profit_ratio REAL,
                our_profit_ratio REAL
            );
            INSERT INTO partners VALUES (1, '波兰合伙人', 'CNY', 0.5, 0.25, 0.25);
            INSERT INTO partners VALUES (2, '澳洲合伙人', 'CNY', 0.5, 0.25, 0.25);
            CREATE TABLE partner_sites (
                partner_id INTEGER,
                site_id INTEGER
            );
            INSERT INTO partner_sites VALUES (1, 1);
            INSERT INTO partner_sites VALUES (2, 2);

            CREATE TABLE orders (
                id TEXT PRIMARY KEY,
                date_created TEXT,
                source TEXT,
                currency TEXT,
                total REAL,
                shipping_total REAL,
                status TEXT,
                payment_method TEXT,
                is_undelivered INTEGER DEFAULT 0,
                is_problem_return INTEGER DEFAULT 0,
                shipping_loss_amount REAL DEFAULT 0,
                product_loss_amount REAL DEFAULT 0
            );
            INSERT INTO orders VALUES (
                'pl-1', '2026-07-10T10:00:00', 'https://pl.example',
                'CNY', 1000, 100, 'completed', 'cod', 0, 0, 0, 0
            );
            INSERT INTO orders VALUES (
                'au-1', '2026-07-11T10:00:00', 'https://au.example',
                'CNY', 120, 20, 'completed', 'stripe', 0, 0, 0, 0
            );

            CREATE TABLE exchange_rates (
                id INTEGER PRIMARY KEY,
                year_month TEXT,
                currency TEXT,
                rate_to_cny REAL
            );
            CREATE TABLE sales_board_exchange_rates (
                id INTEGER PRIMARY KEY,
                year_month TEXT,
                currency TEXT,
                rate_to_cny REAL
            );
            CREATE TABLE sales_targets (
                id INTEGER PRIMARY KEY,
                year_month TEXT,
                manager TEXT,
                base_salary REAL,
                commission_rate REAL
            );
            INSERT INTO sales_targets VALUES (1, '2026-07', '吴辉', 100, 0.05);
            INSERT INTO sales_targets VALUES (2, '2026-07', '付肖肖', 100, 0.05);
            INSERT INTO sales_targets VALUES (3, '2026-07', '零销量员工', 100, 0.05);
            CREATE TABLE sales_groups (
                id INTEGER PRIMARY KEY,
                leader_manager TEXT,
                bonus_rate REAL
            );
            CREATE TABLE sales_group_members (
                id INTEGER PRIMARY KEY,
                group_id INTEGER,
                manager TEXT
            );
            """
        )
        conn.commit()
        conn.close()
        init_company_profit_tables(self.get_conn)

    def tearDown(self):
        self.tempdir.cleanup()

    def get_conn(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    @staticmethod
    def revenue_condition(prefix=""):
        p = f"{prefix}." if prefix else ""
        return f"{p}status = 'completed'"

    @staticmethod
    def board(month):
        if month != "2026-07":
            return {
                "board_data": [],
                "group_summaries": [],
                "team_totals": {
                    "month_net_cny": 0,
                    "total_income": 0,
                    "country_profit": {},
                },
            }
        return {
            "board_data": [
                {
                    "manager": "吴辉",
                    "month_net_cny": 900,
                    "commission_base_cny": 900,
                    "base_salary": 100,
                },
                {
                    "manager": "付肖肖",
                    "month_net_cny": 100,
                    "commission_base_cny": 100,
                    "base_salary": 100,
                },
            ],
            "group_summaries": [],
            "team_totals": {
                "month_net_cny": 1000,
                "total_income": 250,
                "total_commission": 50,
                "country_profit": {
                    "PL": {"net_cny": 900},
                    "AU": {"net_cny": 100},
                },
            },
        }

    @staticmethod
    def partner_detail(partner_id, year, month):
        if (year, month) != (2026, 7):
            return {
                "total_net_pln": 0,
                "actual_cost_pln": 0,
                "shipping_loss": 0,
                "cost_unmapped_qty": 0,
                "cost_unmapped_revenue_pln": 0,
            }
        if partner_id == 1:
            return {
                "total_net_pln": 900,
                "actual_cost_pln": 400,
                "shipping_loss": 0,
                "cost_unmapped_qty": 0,
                "cost_unmapped_revenue_pln": 0,
            }
        return {
            "total_net_pln": 100,
            "actual_cost_pln": 40,
            "shipping_loss": 0,
            "cost_unmapped_qty": 2,
            "cost_unmapped_revenue_pln": 20,
        }

    @staticmethod
    def statement_split(
        net,
        actual_cost,
        cost_ratio,
        partner_ratio,
        our_ratio,
        mode,
        shipping_loss=0,
    ):
        if mode == "actual":
            remainder = net - actual_cost
            ratio_total = partner_ratio + our_ratio
            return (
                actual_cost,
                remainder * partner_ratio / ratio_total,
                remainder * our_ratio / ratio_total,
            )
        successful_net = net + shipping_loss
        return (
            successful_net * cost_ratio,
            successful_net * partner_ratio - shipping_loss / 2,
            successful_net * our_ratio - shipping_loss / 2,
        )

    def test_offline_table_init_is_idempotent_and_adds_no_web_permissions(self):
        conn = self.get_conn()
        initial_columns = {
            row["name"] for row in conn.execute("PRAGMA table_info(users)")
        }
        conn.close()
        self.assertNotIn("can_view_company_profit", initial_columns)
        self.assertNotIn("can_edit_company_profit", initial_columns)

        init_company_profit_tables(self.get_conn)
        conn = self.get_conn()
        second_columns = {
            row["name"] for row in conn.execute("PRAGMA table_info(users)")
        }
        finance_tables = {
            row["name"]
            for row in conn.execute(
                """
                SELECT name FROM sqlite_master
                WHERE type = 'table' AND name LIKE 'company_profit_%'
                """
            )
        }
        conn.close()
        self.assertNotIn("can_view_company_profit", second_columns)
        self.assertNotIn("can_edit_company_profit", second_columns)
        self.assertIn("company_profit_month_settings", finance_tables)
        self.assertIn("company_profit_expenses", finance_tables)

    def test_unknown_market_is_not_guessed_and_actual_can_be_confirmed(self):
        summary = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )
        self.assertEqual(summary["gmv_cny"], 1120)
        self.assertEqual(summary["net_sales_cny"], 1000)
        self.assertEqual(summary["team_net_sales_cny"], 1000)
        self.assertEqual(summary["actual"]["company_revenue_cny"], 225)
        self.assertEqual(summary["actual"]["profit_cny"], -25)
        self.assertFalse(summary["actual"]["complete"])
        self.assertTrue(any("AU" in gap for gap in summary["data_gaps"]))
        self.assertEqual(
            summary["forecast"]["payroll_detail"]["base_salary_cny"], 300
        )

        conn = self.get_conn()
        conn.execute(
            """
            INSERT INTO company_profit_market_rules (
                year_month, country, share_rate
            ) VALUES ('2026-07', 'AU', 0.5)
            """
        )
        conn.execute(
            """
            INSERT INTO company_profit_month_settings (
                year_month, payroll_actual_override, actual_expenses_complete
            ) VALUES ('2026-07', 205, 1)
            """
        )
        conn.execute(
            """
            INSERT INTO company_profit_expenses (
                year_month, scenario, category, name, amount_cny
            ) VALUES ('2026-07', 'actual', '服务器/域名', '服务器', 20)
            """
        )
        conn.commit()
        conn.close()

        confirmed = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )
        self.assertEqual(confirmed["actual"]["company_revenue_cny"], 275)
        self.assertEqual(confirmed["actual"]["profit_cny"], 50)
        self.assertTrue(confirmed["actual"]["complete"])

    def test_each_sales_board_month_is_computed_once_per_summary(self):
        calls = []

        def tracked_board(month):
            calls.append(month)
            return self.board(month)

        build_company_profit_summary(
            self.get_conn,
            tracked_board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )

        self.assertEqual(len(calls), len(set(calls)), calls)
        self.assertEqual(
            set(calls),
            {"2026-07", "2026-06", "2026-05", "2026-04"},
        )

    def test_future_forecast_flags_share_missing_only_in_prediction(self):
        summary = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-08",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )
        self.assertEqual(summary["actual"]["company_revenue_cny"], 0)
        self.assertGreater(summary["forecast"]["company_revenue_cny"], 0)
        self.assertFalse(summary["forecast"]["complete"])
        self.assertTrue(
            any(
                "预测数据缺少市场分润比例" in gap and "AU" in gap
                for gap in summary["data_gaps"]
            )
        )

    def test_partner_reconciliation_switches_between_contract_and_actual_cost(self):
        percentage = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=self.partner_detail,
            statement_split=self.statement_split,
        )
        self.assertEqual(percentage["calculation_mode"], "percentage")
        self.assertEqual(percentage["actual"]["company_revenue_cny"], 250)
        self.assertEqual(
            next(
                row
                for row in percentage["countries"]
                if row["country"] == "PL"
            )["share_source"],
            "合伙人对账 · 约定比例",
        )

        conn = self.get_conn()
        conn.execute(
            """
            INSERT INTO company_profit_month_settings (
                year_month, calculation_mode
            ) VALUES ('2026-07', 'actual_cost')
            """
        )
        conn.commit()
        conn.close()

        actual_cost = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=self.partner_detail,
            statement_split=self.statement_split,
        )
        self.assertEqual(actual_cost["calculation_mode"], "actual_cost")
        self.assertEqual(actual_cost["actual"]["company_revenue_cny"], 280)
        pl = next(
            row
            for row in actual_cost["countries"]
            if row["country"] == "PL"
        )
        self.assertEqual(pl["actual_product_cost_cny"], 400)
        self.assertEqual(pl["share_rate"], 0.5)
        self.assertTrue(pl["share_locked"])
        self.assertFalse(actual_cost["actual"]["complete"])
        self.assertTrue(
            any(
                "AU实际成本未完整匹配：2件" in gap
                for gap in actual_cost["data_gaps"]
            )
        )

    def test_actual_cost_share_uses_configured_ratio_not_rounded_money(self):
        conn = self.get_conn()
        conn.execute(
            """
            INSERT INTO company_profit_month_settings (
                year_month, calculation_mode
            ) VALUES ('2026-07', 'actual_cost')
            """
        )
        conn.commit()
        conn.close()

        def detail(partner_id, year, month):
            if partner_id == 1:
                return {
                    "total_net_pln": 100,
                    "actual_cost_pln": 33.33,
                    "shipping_loss": 0,
                    "cost_unmapped_qty": 0,
                    "cost_unmapped_revenue_pln": 0,
                }
            return {
                "total_net_pln": 100,
                "actual_cost_pln": 40,
                "shipping_loss": 0,
                "cost_unmapped_qty": 0,
                "cost_unmapped_revenue_pln": 0,
            }

        def rounded_split(
            net,
            actual_cost,
            cost_ratio,
            partner_ratio,
            our_ratio,
            mode,
            shipping_loss=0,
        ):
            if mode == "actual":
                remainder = net - actual_cost
                ratio_total = partner_ratio + our_ratio
                return (
                    round(actual_cost, 2),
                    round(remainder * partner_ratio / ratio_total, 2),
                    round(remainder * our_ratio / ratio_total, 2),
                )
            return (
                round(net * cost_ratio, 2),
                round(net * partner_ratio, 2),
                round(net * our_ratio, 2),
            )

        summary = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=detail,
            statement_split=rounded_split,
        )
        pl = next(
            row for row in summary["countries"] if row["country"] == "PL"
        )
        self.assertEqual(pl["share_rate"], 0.5)

    def test_forecast_reuses_prior_actual_cost_and_saved_scenario(self):
        conn = self.get_conn()
        conn.executescript(
            """
            CREATE TABLE reconciliation_statements (
                id INTEGER PRIMARY KEY,
                partner_id INTEGER,
                period_year INTEGER,
                period_month INTEGER,
                total_net_pln REAL,
                actual_cost_pln_snapshot REAL,
                calc_mode TEXT,
                exchange_rate_cny REAL,
                status TEXT,
                is_manual INTEGER DEFAULT 0,
                updated_at TEXT
            );
            INSERT INTO reconciliation_statements VALUES (
                1, 1, 2026, 6, 900, 400, 'actual', 1,
                'generated', 0, '2026-07-01 10:00:00'
            );
            INSERT INTO reconciliation_statements VALUES (
                2, 2, 2026, 6, 100, 40, 'actual', 1,
                'generated', 0, '2026-07-01 10:05:00'
            );
            INSERT INTO company_profit_month_settings (
                year_month, calculation_mode
            ) VALUES ('2026-07', 'actual_cost');
            """
        )
        conn.commit()
        conn.close()

        forecast = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=self.partner_detail,
            statement_split=self.statement_split,
            prefer_reconciled_snapshots=True,
        )
        self.assertEqual(forecast["actual"]["company_revenue_cny"], 0)
        self.assertGreater(forecast["forecast"]["company_revenue_cny"], 0)
        self.assertGreater(forecast["scenario"]["company_revenue_cny"], 0)
        self.assertTrue(
            all(
                row["forecast_revenue_source"]
                == "沿用最近已确认月份实际成本率"
                for row in forecast["countries"]
                if row["country"] in {"PL", "AU"}
            )
        )

        conn = self.get_conn()
        conn.execute(
            """
            INSERT INTO company_profit_forecast_scenarios (
                year_month, target_net_sales_cny, company_revenue_rate,
                payroll_cny, fixed_cost_cny, variable_cost_cny
            ) VALUES ('2026-07', 800000, 0.28, 85182.18, 11074, 0)
            """
        )
        conn.commit()
        conn.close()

        saved = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=self.partner_detail,
            statement_split=self.statement_split,
            prefer_reconciled_snapshots=True,
        )
        self.assertTrue(saved["scenario"]["is_saved"])
        self.assertEqual(saved["scenario"]["company_revenue_cny"], 224000)
        self.assertEqual(saved["scenario"]["profit_cny"], 127743.82)
        ladder = saved["revenue_ladder"]
        self.assertEqual(
            [row["label"] for row in ladder["rows"][:4]],
            [
                "7月截至当前",
                "7月月底预测",
                "7月月底预测增长20%",
                "7月⬆️10万",
            ],
        )
        self.assertEqual(ladder["base_net_sales_cny"], 2066.67)
        self.assertEqual(ladder["rows"][0]["target_net_sales_cny"], 1000)
        self.assertEqual(ladder["rows"][0]["company_revenue_cny"], 280)
        self.assertEqual(ladder["rows"][0]["payroll_cny"], 250)
        self.assertEqual(ladder["rows"][0]["profit_cny"], 30)
        self.assertEqual(ladder["payroll_mode"], "full_forecast")

    def test_revenue_ladder_rounds_june_targets_to_5_and_10_wan_steps(self):
        ladder = _revenue_ladder(
            "2026-06",
            367875.5,
            0.2823580578404445,
            367357.44,
            {
                "base_salary_cny": 34000,
                "commission_cny": 18124.48,
                "leader_bonus_cny": 4577.81,
            },
            None,
            25074,
            "当月实际支出",
        )
        self.assertEqual(
            [row["target_net_sales_cny"] for row in ladder["rows"][:7]],
            [
                367875.5,
                441450.6,
                450000,
                500000,
                550000,
                600000,
                700000,
            ],
        )
        self.assertEqual(
            [row["label"] for row in ladder["rows"][:3]],
            ["6月当月", "6月当月增长20%", "6月⬆️45万"],
        )
        self.assertEqual(ladder["rows"][-1]["target_net_sales_cny"], 1400000)
        self.assertEqual(ladder["daily_operations_cny"], 25074)

    def test_historical_ladder_anchors_actual_payroll_and_only_scales_growth(self):
        ladder = _revenue_ladder(
            "2026-06",
            367875.5,
            0.2823580578404445,
            367357.44,
            {
                "base_salary_cny": 34000,
                "commission_cny": 18124.48,
                "leader_bonus_cny": 4577.81,
            },
            None,
            25074,
            "当月实际支出",
            52875.88,
            "实际已发（手工确认）",
        )
        self.assertEqual(ladder["rows"][0]["payroll_cny"], 52875.88)
        self.assertEqual(ladder["rows"][1]["payroll_cny"], 57422.74)
        self.assertEqual(
            ladder["payroll_method"],
            "实际已发（手工确认） + 仅新增销售额对应的提成与带团奖金",
        )
        self.assertEqual(ladder["breakeven_net_sales_cny"], 276067.49)

    def test_current_month_ladder_separates_current_snapshot_and_month_end_forecast(self):
        ladder = _revenue_ladder(
            "2026-07",
            768627.32,
            0.278374,
            768627.32,
            {
                "base_salary_cny": 36000,
                "commission_cny": 38088.88,
                "leader_bonus_cny": 11271.55,
            },
            None,
            25074,
            "当月预测与周期性继承支出",
            current_snapshot_net_sales_cny=694244.03,
            current_snapshot_payroll_cny=80583.61,
            current_snapshot_payroll_source="销售看板动态暂估",
        )

        self.assertEqual(
            [row["label"] for row in ladder["rows"][:3]],
            ["7月截至当前", "7月月底预测", "7月月底预测增长20%"],
        )
        self.assertEqual(ladder["rows"][0]["target_net_sales_cny"], 694244.03)
        self.assertEqual(ladder["rows"][0]["payroll_cny"], 80583.61)
        self.assertEqual(ladder["rows"][1]["target_net_sales_cny"], 768627.32)
        self.assertEqual(ladder["rows"][1]["payroll_cny"], 85360.43)
        self.assertEqual(ladder["base_net_sales_cny"], 768627.32)
        self.assertEqual(ladder["payroll_mode"], "full_forecast")

    def test_historical_month_prefers_its_actual_operating_expenses(self):
        conn = self.get_conn()
        conn.executescript(
            """
            INSERT INTO company_profit_expenses (
                year_month, scenario, category, name, amount_cny,
                is_recurring
            ) VALUES (
                '2026-06', 'actual', '服务器/域名', '日常运营',
                100, 1
            );
            INSERT INTO company_profit_expenses (
                year_month, scenario, category, name, amount_cny,
                is_recurring
            ) VALUES (
                '2026-06', 'forecast', '其他', '预测费用',
                60, 0
            );
            """
        )
        conn.commit()
        conn.close()

        june = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-06",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )
        self.assertEqual(june["actual"]["other_expenses_cny"], 100)
        self.assertEqual(june["forecast"]["other_expenses_cny"], 100)
        self.assertEqual(
            june["forecast"]["other_expenses_source"],
            "当月实际支出",
        )
        self.assertEqual(
            june["revenue_ladder"]["daily_operations_cny"],
            100,
        )
        self.assertEqual(
            june["revenue_ladder"]["daily_operations_source"],
            "当月实际支出",
        )

    def test_recurring_expense_can_be_overridden_for_one_month_only(self):
        conn = self.get_conn()
        conn.execute(
            """
            INSERT INTO company_profit_expenses (
                year_month, scenario, category, name, amount_cny,
                is_recurring
            ) VALUES (
                '2026-06', 'actual', '服务器/域名', '日常运营',
                100, 1
            )
            """
        )
        conn.commit()
        conn.close()

        inherited = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )
        self.assertEqual(
            inherited["revenue_ladder"]["daily_operations_cny"],
            100,
        )
        self.assertEqual(
            inherited["recurring_forecast_expenses"][0]["amount_cny"],
            100,
        )

        conn = self.get_conn()
        conn.execute(
            """
            INSERT INTO company_profit_expenses (
                year_month, scenario, category, name, amount_cny,
                is_recurring, notes
            ) VALUES (
                '2026-07', 'forecast', '服务器/域名', '日常运营',
                60, 0, '本月覆盖'
            )
            """
        )
        conn.commit()
        conn.close()

        july = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )
        self.assertEqual(july["recurring_forecast_expenses"], [])
        self.assertEqual(
            july["revenue_ladder"]["daily_operations_cny"],
            60,
        )

        august = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-08",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
        )
        self.assertEqual(
            august["revenue_ladder"]["daily_operations_cny"],
            100,
        )

    def test_reconciled_snapshot_mode_fails_closed_when_snapshot_is_missing(self):
        conn = self.get_conn()
        conn.execute(
            """
            INSERT INTO company_profit_month_settings (
                year_month, calculation_mode
            ) VALUES ('2026-07', 'actual_cost')
            """
        )
        conn.commit()
        conn.close()

        summary = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=self.partner_detail,
            statement_split=self.statement_split,
            prefer_reconciled_snapshots=True,
        )
        self.assertEqual(summary["actual"]["company_revenue_cny"], 0)
        self.assertFalse(summary["actual"]["complete"])
        self.assertTrue(
            any(
                "合伙人对账单实际成本快照" in gap
                and "PL" in gap
                and "AU" in gap
                for gap in summary["data_gaps"]
            )
        )

    def test_reconciled_snapshot_is_used_as_actual_cost_source(self):
        conn = self.get_conn()
        conn.executescript(
            """
            CREATE TABLE reconciliation_statements (
                id INTEGER PRIMARY KEY,
                partner_id INTEGER,
                period_year INTEGER,
                period_month INTEGER,
                total_net_pln REAL,
                actual_cost_pln_snapshot REAL,
                calc_mode TEXT,
                exchange_rate_cny REAL,
                status TEXT,
                is_manual INTEGER DEFAULT 0,
                updated_at TEXT
            );
            CREATE TABLE reconciliation_statement_orders (
                statement_id INTEGER,
                shipping_loss_at_gen REAL,
                is_undelivered_at_gen INTEGER
            );
            INSERT INTO reconciliation_statements VALUES (
                1, 1, 2026, 7, 900, 400, 'actual', 2,
                'generated', 0, '2026-07-20 10:00:00'
            );
            INSERT INTO reconciliation_statements VALUES (
                2, 2, 2026, 7, 100, 40, 'actual', 3,
                'generated', 0, '2026-07-20 10:05:00'
            );
            INSERT INTO reconciliation_statement_orders VALUES (1, 10, 1);
            INSERT INTO reconciliation_statement_orders VALUES (2, 0, 0);
            UPDATE orders
            SET is_undelivered = 1, shipping_loss_amount = 999
            WHERE id = 'pl-1';
            INSERT INTO company_profit_month_settings (
                year_month, calculation_mode
            ) VALUES ('2026-07', 'actual_cost');
            """
        )
        conn.commit()
        conn.close()

        summary = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=self.partner_detail,
            statement_split=self.statement_split,
            prefer_reconciled_snapshots=True,
        )
        self.assertEqual(summary["net_sales_cny"], 2100)
        self.assertEqual(summary["team_net_sales_cny"], 1000)
        self.assertEqual(summary["actual"]["company_revenue_cny"], 590)
        pl = next(
            row for row in summary["countries"] if row["country"] == "PL"
        )
        self.assertEqual(pl["actual_product_cost_cny"], 800)
        self.assertEqual(
            pl["reconciliation_sources"],
            ["合伙人对账单快照"],
        )
        self.assertEqual(
            pl["statement_updated_at"],
            ["2026-07-20 10:00:00"],
        )

        conn = self.get_conn()
        conn.execute(
            """
            UPDATE company_profit_month_settings
            SET calculation_mode = 'percentage'
            WHERE year_month = '2026-07'
            """
        )
        conn.commit()
        conn.close()
        percentage = build_company_profit_summary(
            self.get_conn,
            self.board,
            self.revenue_condition,
            "2026-07",
            today=__import__("datetime").date(2026, 7, 15),
            trend_count=2,
            partner_recon_detail=self.partner_detail,
            statement_split=self.statement_split,
            prefer_reconciled_snapshots=True,
        )
        self.assertEqual(
            percentage["actual"]["company_revenue_cny"],
            520,
        )


if __name__ == "__main__":
    unittest.main()
