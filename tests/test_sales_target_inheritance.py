import sqlite3

from sales_target_inheritance import load_sales_targets_for_month


def make_conn():
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE sales_targets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year_month TEXT NOT NULL,
            manager TEXT NOT NULL,
            monthly_target REAL DEFAULT 0,
            weekly_targets TEXT DEFAULT '{}',
            base_salary REAL DEFAULT 7000,
            commission_rate REAL DEFAULT 0.05,
            notes TEXT DEFAULT '',
            created_at TEXT,
            updated_at TEXT,
            UNIQUE(year_month, manager)
        )
        """
    )
    return conn


def insert_target(
    conn,
    month,
    manager,
    monthly_target,
    base_salary,
    commission_rate,
    weekly_targets='{"w1": 100}',
    notes="旧月份备注",
):
    conn.execute(
        """
        INSERT INTO sales_targets (
            year_month, manager, monthly_target, weekly_targets,
            base_salary, commission_rate, notes
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            month,
            manager,
            monthly_target,
            weekly_targets,
            base_salary,
            commission_rate,
            notes,
        ),
    )


def test_new_month_inherits_latest_salary_rules_only():
    conn = make_conn()
    insert_target(conn, "2026-05", "吴辉", 400000, 9000, 0.04)
    insert_target(conn, "2026-06", "吴辉", 500000, 10000, 0.05)

    july = load_sales_targets_for_month(conn, "2026-07")
    inherited = july["吴辉"]

    assert inherited["year_month"] == "2026-07"
    assert inherited["monthly_target"] == 0
    assert inherited["weekly_targets"] == "{}"
    assert inherited["base_salary"] == 10000
    assert inherited["commission_rate"] == 0.05
    assert inherited["notes"] == ""
    assert inherited["salary_inherited_from_month"] == "2026-06"


def test_exact_month_target_wins_and_explicit_zero_salary_is_preserved():
    conn = make_conn()
    insert_target(conn, "2026-06", "刘皓玮", 300000, 7000, 0.05)
    insert_target(
        conn,
        "2026-07",
        "刘皓玮",
        350000,
        0,
        0.06,
        weekly_targets='{"w1": 80000}',
        notes="当月设置",
    )

    july = load_sales_targets_for_month(conn, "2026-07")
    exact = july["刘皓玮"]

    assert exact["monthly_target"] == 350000
    assert exact["weekly_targets"] == '{"w1": 80000}'
    assert exact["base_salary"] == 0
    assert exact["commission_rate"] == 0.06
    assert exact["notes"] == "当月设置"
    assert exact["salary_inherited_from_month"] is None


def test_future_target_is_never_used_as_an_inheritance_source():
    conn = make_conn()
    insert_target(conn, "2026-06", "王松", 300000, 8000, 0.05)
    insert_target(conn, "2026-08", "王松", 600000, 12000, 0.08)

    july = load_sales_targets_for_month(conn, "2026-07")

    assert july["王松"]["base_salary"] == 8000
    assert july["王松"]["commission_rate"] == 0.05
    assert july["王松"]["salary_inherited_from_month"] == "2026-06"
