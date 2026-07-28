"""Company profitability dashboard.

The sales board measures sales performance and estimated compensation.  This
module keeps company economics separate: recognized company revenue, payroll,
operating expenses, provisional actual profit, and a transparent forecast.
"""

from calendar import monthrange
from collections import defaultdict
from datetime import date
import json
import math
import re


MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
DEFAULT_SHARE_RATES = {"PL": 0.25}
EXPENSE_SCENARIOS = {"actual", "forecast"}
CALCULATION_MODES = {"percentage", "actual_cost"}
MAX_AMOUNT_CNY = 1_000_000_000
FORECAST_SALES_GROWTH_RATE = 0.20
FORECAST_SALES_MIN_ADDITION_CNY = 50_000
FORECAST_SALES_FINE_STEP_CNY = 50_000
FORECAST_SALES_FINE_LIMIT_CNY = 600_000
FORECAST_SALES_COARSE_STEP_CNY = 100_000
FORECAST_SALES_MAX_ADDITION_CNY = 1_000_000


def _month_shift(year_month, delta):
    year, month = (int(part) for part in year_month.split("-"))
    absolute = year * 12 + month - 1 + delta
    return f"{absolute // 12}-{absolute % 12 + 1:02d}"


def _month_sequence(end_month, count):
    return [_month_shift(end_month, offset) for offset in range(1 - count, 1)]


def _valid_month(value):
    value = (value or "").strip()
    if not MONTH_RE.match(value):
        raise ValueError("月份格式必须为 YYYY-MM")
    return value


def _optional_amount(value, field_name):
    if value in (None, ""):
        return None
    try:
        amount = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{field_name}必须是数字")
    if not math.isfinite(amount) or amount < 0 or amount > MAX_AMOUNT_CNY:
        raise ValueError(f"{field_name}超出允许范围")
    return round(amount, 2)


def _required_amount(value, field_name):
    amount = _optional_amount(value, field_name)
    if amount is None:
        raise ValueError(f"{field_name}不能为空")
    return amount


def _table_columns(conn, table):
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}


def _partner_sites_relation(conn):
    """Use live country inheritance when available; keep test/legacy fallback."""
    row = conn.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE name = 'effective_partner_sites'
          AND type IN ('view', 'table')
        """
    ).fetchone()
    return "effective_partner_sites" if row else "partner_sites"


def init_company_profit_tables(get_db_connection):
    """Create finance tables on an offline snapshot database."""
    conn = get_db_connection()
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS company_profit_month_settings (
                year_month TEXT PRIMARY KEY,
                calculation_mode TEXT NOT NULL DEFAULT 'percentage',
                payroll_actual_override REAL,
                payroll_forecast_override REAL,
                actual_expenses_complete INTEGER DEFAULT 0,
                notes TEXT DEFAULT '',
                updated_by TEXT DEFAULT '',
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        settings_columns = _table_columns(conn, "company_profit_month_settings")
        if "calculation_mode" not in settings_columns:
            conn.execute(
                """
                ALTER TABLE company_profit_month_settings
                ADD COLUMN calculation_mode TEXT NOT NULL DEFAULT 'percentage'
                """
            )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS company_profit_market_rules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year_month TEXT NOT NULL,
                country TEXT NOT NULL,
                share_rate REAL,
                forecast_net_sales_cny REAL,
                notes TEXT DEFAULT '',
                updated_by TEXT DEFAULT '',
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(year_month, country)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS company_profit_expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year_month TEXT NOT NULL,
                scenario TEXT NOT NULL CHECK (scenario IN ('actual', 'forecast')),
                category TEXT NOT NULL,
                name TEXT NOT NULL,
                amount_cny REAL NOT NULL CHECK (amount_cny >= 0),
                is_recurring INTEGER DEFAULT 0,
                notes TEXT DEFAULT '',
                created_by TEXT DEFAULT '',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS company_profit_forecast_scenarios (
                year_month TEXT PRIMARY KEY,
                target_net_sales_cny REAL NOT NULL,
                company_revenue_rate REAL NOT NULL,
                payroll_cny REAL NOT NULL,
                fixed_cost_cny REAL NOT NULL,
                variable_cost_cny REAL NOT NULL DEFAULT 0,
                notes TEXT DEFAULT '',
                updated_by TEXT DEFAULT '',
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_company_profit_expense_month
            ON company_profit_expenses(year_month, scenario)
            """
        )
        conn.commit()
    finally:
        conn.close()
def _rate_to_cny(conn, currency, year_month):
    currency = (currency or "CNY").upper()
    if currency == "CNY":
        return 1.0
    try:
        row = conn.execute(
            """
            SELECT rate_to_cny
            FROM sales_board_exchange_rates
            WHERE currency = ? AND year_month = ?
            """,
            (currency, year_month),
        ).fetchone()
    except sqlite3.OperationalError:
        row = None
    if row:
        return float(row["rate_to_cny"])

    row = conn.execute(
        """
        SELECT rate_to_cny
        FROM exchange_rates
        WHERE currency = ? AND year_month <= ?
        ORDER BY year_month DESC
        LIMIT 1
        """,
        (currency, year_month),
    ).fetchone()
    if not row:
        row = conn.execute(
            """
            SELECT rate_to_cny
            FROM exchange_rates
            WHERE currency = ?
            ORDER BY year_month DESC
            LIMIT 1
            """,
            (currency,),
        ).fetchone()
    return float(row["rate_to_cny"]) if row else None


def _month_gmv(conn, year_month):
    total = 0.0
    by_country = defaultdict(float)
    missing_rates = set()
    rows = conn.execute(
        """
        SELECT COALESCE(s.country, '未分配') AS country,
               o.currency,
               SUM(o.total) AS amount
        FROM orders o
        LEFT JOIN sites s ON s.url = o.source
        WHERE strftime('%Y-%m', o.date_created) = ?
          AND o.status NOT IN ('checkout-draft', 'trash')
        GROUP BY country, o.currency
        """,
        (year_month,),
    ).fetchall()
    for row in rows:
        rate = _rate_to_cny(conn, row["currency"], year_month)
        if rate is None:
            missing_rates.add((row["currency"] or "N/A").upper())
            continue
        amount = float(row["amount"] or 0) * rate
        total += amount
        by_country[row["country"]] += amount
    return round(total, 2), {
        key: round(value, 2) for key, value in by_country.items()
    }, sorted(missing_rates)


def _month_settings(conn, year_month):
    row = conn.execute(
        """
        SELECT *
        FROM company_profit_month_settings
        WHERE year_month = ?
        """,
        (year_month,),
    ).fetchone()
    if not row:
        return {
            "year_month": year_month,
            "calculation_mode": "percentage",
            "payroll_actual_override": None,
            "payroll_forecast_override": None,
            "actual_expenses_complete": 0,
            "notes": "",
        }
    result = dict(row)
    if result.get("calculation_mode") not in CALCULATION_MODES:
        result["calculation_mode"] = "percentage"
    return result


def _market_rules(conn, year_month):
    rows = conn.execute(
        """
        SELECT country, share_rate, forecast_net_sales_cny, notes
        FROM company_profit_market_rules
        WHERE year_month = ?
        ORDER BY country
        """,
        (year_month,),
    ).fetchall()
    return {row["country"]: dict(row) for row in rows}


def _forecast_scenario(
    conn,
    year_month,
    default_net_sales_cny,
    default_company_revenue_rate,
    default_payroll_cny,
    default_fixed_cost_cny,
):
    row = conn.execute(
        """
        SELECT year_month, target_net_sales_cny, company_revenue_rate,
               payroll_cny, fixed_cost_cny, variable_cost_cny, notes,
               updated_by, updated_at
        FROM company_profit_forecast_scenarios
        WHERE year_month = ?
        """,
        (year_month,),
    ).fetchone()
    saved = dict(row) if row else {}

    target_net_sales = float(
        saved.get("target_net_sales_cny", default_net_sales_cny) or 0
    )
    company_revenue_rate = float(
        saved.get(
            "company_revenue_rate",
            default_company_revenue_rate,
        )
        or 0
    )
    payroll = float(saved.get("payroll_cny", default_payroll_cny) or 0)
    fixed_cost = float(
        saved.get("fixed_cost_cny", default_fixed_cost_cny) or 0
    )
    variable_cost = float(saved.get("variable_cost_cny", 0) or 0)
    company_revenue = target_net_sales * company_revenue_rate
    total_cost = payroll + fixed_cost + variable_cost
    profit = company_revenue - total_cost
    return {
        "target_net_sales_cny": round(target_net_sales, 2),
        "company_revenue_rate": company_revenue_rate,
        "company_revenue_cny": round(company_revenue, 2),
        "payroll_cny": round(payroll, 2),
        "fixed_cost_cny": round(fixed_cost, 2),
        "variable_cost_cny": round(variable_cost, 2),
        "total_cost_cny": round(total_cost, 2),
        "profit_cny": round(profit, 2),
        "margin_pct": (
            round(profit / company_revenue * 100, 1)
            if abs(company_revenue) > 0.01
            else None
        ),
        "breakeven_net_sales_cny": (
            round(total_cost / company_revenue_rate, 2)
            if company_revenue_rate > 0
            else None
        ),
        "is_saved": bool(row),
        "notes": saved.get("notes", ""),
        "updated_by": saved.get("updated_by", ""),
        "updated_at": saved.get("updated_at"),
        "target_source": (
            "手工情景"
            if row
            else "各市场本月预测净销售汇总"
        ),
        "rate_source": (
            "手工情景"
            if row
            else "各市场预测收入率加权"
        ),
        "payroll_source": (
            "手工情景"
            if row
            else "系统预测工资"
        ),
        "fixed_cost_source": (
            "手工情景"
            if row
            else "本月预测与周期性支出"
        ),
    }


def _revenue_ladder(
    year_month,
    base_net_sales_cny,
    company_revenue_rate,
    forecast_net_sales_cny,
    forecast_payroll_detail,
    forecast_payroll_override,
    daily_operations_cny,
    daily_operations_source,
    payroll_anchor_cny=None,
    payroll_anchor_source=None,
    current_snapshot_net_sales_cny=None,
    current_snapshot_payroll_cny=None,
    current_snapshot_payroll_source=None,
):
    """Build comparable sales stages using one transparent cost formula."""
    base_net_sales = float(base_net_sales_cny or 0)
    company_rate = float(company_revenue_rate or 0)
    daily_operations = float(daily_operations_cny or 0)

    forecast_variable_payroll = (
        float(forecast_payroll_detail.get("commission_cny") or 0)
        + float(forecast_payroll_detail.get("leader_bonus_cny") or 0)
    )
    variable_payroll_rate = (
        forecast_variable_payroll / float(forecast_net_sales_cny)
        if abs(float(forecast_net_sales_cny or 0)) > 0.01
        else 0.0
    )

    if payroll_anchor_cny is not None:
        fixed_payroll = float(payroll_anchor_cny or 0)
        payroll_mode = "actual_anchor_incremental"
        payroll_method = (
            f"{payroll_anchor_source or '当月实际工资'}"
            " + 仅新增销售额对应的提成与带团奖金"
        )
    elif forecast_payroll_override is not None:
        fixed_payroll = float(forecast_payroll_override or 0)
        variable_payroll_rate = 0.0
        payroll_mode = "forecast_override"
        payroll_method = "手工预测工资固定值"
    else:
        fixed_payroll = float(
            forecast_payroll_detail.get("base_salary_cny") or 0
        )
        payroll_mode = "full_forecast"
        payroll_method = "全额底薪 + 随销售额同比例变化的提成与带团奖金"

    has_current_snapshot = (
        current_snapshot_net_sales_cny is not None
        and current_snapshot_payroll_cny is not None
    )
    if has_current_snapshot:
        payroll_method = (
            f"{current_snapshot_payroll_source or '销售看板动态暂估'}（截至当前）；"
            "月底预测及增长档位按全额底薪 + "
            "随销售额同比例变化的提成与带团奖金"
        )

    def rounded_target(value):
        step = (
            FORECAST_SALES_FINE_STEP_CNY
            if value <= FORECAST_SALES_FINE_LIMIT_CNY
            else FORECAST_SALES_COARSE_STEP_CNY
        )
        return int(math.ceil(float(value) / step) * step)

    month_label = f"{int(year_month.split('-')[1])}月"
    growth_target = base_net_sales * (1 + FORECAST_SALES_GROWTH_RATE)
    first_rounded_target = rounded_target(
        max(
            base_net_sales + FORECAST_SALES_MIN_ADDITION_CNY,
            growth_target,
        )
    )
    last_rounded_target = rounded_target(
        max(
            base_net_sales + FORECAST_SALES_MAX_ADDITION_CNY,
            growth_target,
        )
    )
    if has_current_snapshot:
        sales_targets = [
            (
                f"{month_label}截至当前",
                float(current_snapshot_net_sales_cny or 0),
                float(current_snapshot_payroll_cny or 0),
            ),
            (f"{month_label}月底预测", base_net_sales, None),
            (f"{month_label}月底预测增长20%", growth_target, None),
        ]
    else:
        sales_targets = [
            (f"{month_label}当月", base_net_sales, None),
            (f"{month_label}当月增长20%", growth_target, None),
        ]
    rounded_sales_target = first_rounded_target
    while rounded_sales_target <= last_rounded_target:
        if (
            abs(rounded_sales_target - base_net_sales) > 0.01
            and abs(rounded_sales_target - growth_target) > 0.01
        ):
            sales_targets.append(
                (
                    f"{month_label}⬆️{rounded_sales_target // 10_000}万",
                    rounded_sales_target,
                    None,
                )
            )
        rounded_sales_target += (
            FORECAST_SALES_FINE_STEP_CNY
            if rounded_sales_target < FORECAST_SALES_FINE_LIMIT_CNY
            else FORECAST_SALES_COARSE_STEP_CNY
        )

    rows = []
    for label, target_net_sales, row_payroll_override in sales_targets:
        company_revenue = target_net_sales * company_rate
        if row_payroll_override is not None:
            payroll = row_payroll_override
        elif payroll_mode == "actual_anchor_incremental":
            payroll = fixed_payroll + max(
                target_net_sales - base_net_sales,
                0,
            ) * variable_payroll_rate
        else:
            payroll = fixed_payroll + target_net_sales * variable_payroll_rate
        total_expenses = payroll + daily_operations
        profit = company_revenue - total_expenses
        rows.append(
            {
                "label": label,
                "increment_cny": round(target_net_sales - base_net_sales, 2),
                "target_net_sales_cny": round(target_net_sales, 2),
                "company_revenue_cny": round(company_revenue, 2),
                "payroll_cny": round(payroll, 2),
                "daily_operations_cny": round(daily_operations, 2),
                "total_expenses_cny": round(total_expenses, 2),
                "profit_cny": round(profit, 2),
            }
        )

    if payroll_mode == "actual_anchor_incremental":
        breakeven_net_sales = (
            (fixed_payroll + daily_operations) / company_rate
            if company_rate > 0
            else None
        )
    else:
        breakeven_denominator = company_rate - variable_payroll_rate
        breakeven_net_sales = (
            (fixed_payroll + daily_operations) / breakeven_denominator
            if breakeven_denominator > 0
            else None
        )
    return {
        "base_net_sales_cny": round(base_net_sales, 2),
        "company_revenue_rate": company_rate,
        "fixed_payroll_cny": round(fixed_payroll, 2),
        "variable_payroll_rate": variable_payroll_rate,
        "payroll_mode": payroll_mode,
        "payroll_anchor_cny": (
            round(float(payroll_anchor_cny), 2)
            if payroll_anchor_cny is not None
            else None
        ),
        "payroll_anchor_source": payroll_anchor_source,
        "current_snapshot_net_sales_cny": (
            round(float(current_snapshot_net_sales_cny), 2)
            if current_snapshot_net_sales_cny is not None
            else None
        ),
        "current_snapshot_payroll_cny": (
            round(float(current_snapshot_payroll_cny), 2)
            if current_snapshot_payroll_cny is not None
            else None
        ),
        "daily_operations_cny": round(daily_operations, 2),
        "daily_operations_source": daily_operations_source,
        "payroll_method": payroll_method,
        "breakeven_net_sales_cny": (
            round(breakeven_net_sales, 2)
            if breakeven_net_sales is not None
            else None
        ),
        "rows": rows,
    }


def _partner_reconciliation_basis(
    conn,
    year_month,
    partner_recon_detail,
    statement_split,
    fallback_country_basis=None,
    fallback_has_actual_cost=False,
):
    """Reuse partner-reconciliation actual costs and split formulas by market.

    The reconciliation engine remains the single source of truth for dated
    product-cost matching, unmapped-product warnings, shipping-loss handling,
    and the contract/actual settlement formulas.
    """
    if not partner_recon_detail or not statement_split:
        return {}, []

    year, month = (int(part) for part in year_month.split("-"))
    statement_by_partner = {}
    statement_columns = _table_columns(conn, "reconciliation_statements")
    if {
        "partner_id",
        "period_year",
        "period_month",
        "total_net_pln",
        "actual_cost_pln_snapshot",
        "exchange_rate_cny",
        "is_manual",
    }.issubset(statement_columns):
        statement_by_partner = {
            row["partner_id"]: dict(row)
            for row in conn.execute(
                """
                SELECT id, partner_id, total_net_pln,
                       actual_cost_pln_snapshot, calc_mode,
                       exchange_rate_cny, status, is_manual, updated_at
                FROM reconciliation_statements
                WHERE period_year = ?
                  AND period_month = ?
                """,
                (year, month),
            ).fetchall()
        }

    partner_sites_relation = _partner_sites_relation(conn)
    partner_rows = conn.execute(
        f"""
        SELECT p.id, p.name, p.currency, p.cost_ratio,
               p.partner_profit_ratio, p.our_profit_ratio,
               GROUP_CONCAT(DISTINCT s.country) AS countries,
               COUNT(DISTINCT ps.site_id) AS site_count
        FROM partners p
        LEFT JOIN {partner_sites_relation} ps ON ps.partner_id = p.id
        LEFT JOIN sites s ON s.id = ps.site_id
        GROUP BY p.id
        ORDER BY p.id
        """
    ).fetchall()

    basis = {}
    gaps = []
    for partner in partner_rows:
        countries = sorted(
            {
                value.strip().upper()
                for value in (partner["countries"] or "").split(",")
                if value.strip()
            }
        )
        if not countries or not partner["site_count"]:
            continue
        if len(countries) != 1:
            gaps.append(
                f'{partner["name"]}同时绑定多个市场，实际成本无法自动归属：'
                + "、".join(countries)
            )
            continue

        country = countries[0]
        snapshot = statement_by_partner.get(partner["id"])
        use_snapshot = bool(
            snapshot
            and not snapshot.get("is_manual")
        )
        if use_snapshot:
            detail_is_cny = False
            actual_cost_available = (
                snapshot.get("actual_cost_pln_snapshot") is not None
            )
            snapshot_rate = snapshot.get("exchange_rate_cny")
            if _table_columns(conn, "reconciliation_statement_orders"):
                shipping_loss_row = conn.execute(
                    """
                    SELECT COALESCE(SUM(
                        CASE WHEN COALESCE(is_undelivered_at_gen, 0) = 1
                             THEN COALESCE(shipping_loss_at_gen, 0)
                             ELSE 0 END
                    ), 0) AS shipping_loss
                    FROM reconciliation_statement_orders
                    WHERE statement_id = ?
                    """,
                    (snapshot["id"],),
                ).fetchone()
                snapshot_shipping_loss = shipping_loss_row["shipping_loss"]
            else:
                snapshot_shipping_loss = 0
            detail = {
                "total_net_pln": snapshot["total_net_pln"],
                "actual_cost_pln": snapshot[
                    "actual_cost_pln_snapshot"
                ] or 0,
                "shipping_loss": snapshot_shipping_loss,
                "cost_unmapped_qty": 0,
                "cost_unmapped_revenue_pln": 0,
            }
            basis_source = "合伙人对账单快照"
        else:
            fallback = (fallback_country_basis or {}).get(country)
            if fallback is not None:
                detail_is_cny = True
                actual_cost_available = fallback_has_actual_cost
                detail = {
                    "total_net_pln": fallback.get("net_cny") or 0,
                    "actual_cost_pln": (
                        fallback.get("cost_cny") or 0
                    ),
                    "shipping_loss": (
                        fallback.get("shipping_loss_cny") or 0
                    ),
                    "cost_unmapped_qty": (
                        fallback.get("unmapped_qty")
                        or fallback.get("unmapped_count")
                        or 0
                    ),
                    "cost_unmapped_revenue_pln": (
                        fallback.get("unmapped_revenue_cny") or 0
                    ),
                }
                basis_source = (
                    "销售看板实际成本引擎实时试算"
                    if fallback_has_actual_cost
                    else "销售看板净额实时试算"
                )
            else:
                order_count = conn.execute(
                    f"""
                    SELECT COUNT(*) AS total
                    FROM orders o
                    JOIN sites s ON s.url = o.source
                    JOIN {partner_sites_relation} ps ON ps.site_id = s.id
                    WHERE ps.partner_id = ?
                      AND o.currency = ?
                      AND strftime('%Y-%m', o.date_created) = ?
                    """,
                    (
                        partner["id"],
                        partner["currency"],
                        year_month,
                    ),
                ).fetchone()["total"]
                if order_count:
                    detail = partner_recon_detail(
                        partner["id"], year, month
                    )
                    if not detail:
                        gaps.append(
                            f'{partner["name"]}的合伙人对账数据无法计算'
                        )
                        continue
                    detail_is_cny = False
                    actual_cost_available = True
                    basis_source = "合伙人对账实时计算"
                else:
                    detail = {
                        "total_net_pln": 0,
                        "actual_cost_pln": 0,
                        "shipping_loss": 0,
                        "cost_unmapped_qty": 0,
                        "cost_unmapped_revenue_pln": 0,
                    }
                    detail_is_cny = False
                    actual_cost_available = True
                    basis_source = "本月无合伙人订单"

        rate = (
            1.0
            if detail_is_cny
            else (
                float(snapshot_rate)
                if use_snapshot and snapshot_rate
                else _rate_to_cny(
                    conn, partner["currency"], year_month
                )
            )
        )
        if rate is None:
            gaps.append(
                f'{partner["name"]}缺少{partner["currency"] or "本币"}兑人民币汇率'
            )
            continue

        net_native = float(detail.get("total_net_pln") or 0)
        actual_cost_native = float(detail.get("actual_cost_pln") or 0)
        shipping_loss_native = float(detail.get("shipping_loss") or 0)
        contract_split = statement_split(
            net_native,
            actual_cost_native,
            partner["cost_ratio"],
            partner["partner_profit_ratio"],
            partner["our_profit_ratio"],
            "contract",
            shipping_loss=shipping_loss_native,
        )
        actual_split = statement_split(
            net_native,
            actual_cost_native,
            partner["cost_ratio"],
            partner["partner_profit_ratio"],
            partner["our_profit_ratio"],
            "actual",
            shipping_loss=shipping_loss_native,
        )

        row = basis.setdefault(
            country,
            {
                "partner_names": [],
                "partner_ids": [],
                "basis_sources": [],
                "statement_statuses": [],
                "statement_updated_at": [],
                "actual_cost_available": False,
                "reconciliation_net_cny": 0.0,
                "actual_cost_cny": 0.0,
                "actual_margin_cny": 0.0,
                "contract_company_revenue_cny": 0.0,
                "actual_company_revenue_cny": 0.0,
                "unmapped_qty": 0,
                "unmapped_revenue_cny": 0.0,
                "site_count": 0,
                "configured_margin_share_rates": [],
            },
        )
        net_cny = net_native * rate
        actual_cost_cny = actual_cost_native * rate
        row["partner_names"].append(partner["name"])
        row["partner_ids"].append(partner["id"])
        row["basis_sources"].append(basis_source)
        row["actual_cost_available"] = bool(
            row["actual_cost_available"] or actual_cost_available
        )
        if use_snapshot:
            row["statement_statuses"].append(snapshot.get("status") or "")
            row["statement_updated_at"].append(
                snapshot.get("updated_at") or ""
            )
        row["reconciliation_net_cny"] += net_cny
        row["actual_cost_cny"] += actual_cost_cny
        row["actual_margin_cny"] += (net_native - actual_cost_native) * rate
        row["contract_company_revenue_cny"] += float(contract_split[2]) * rate
        row["actual_company_revenue_cny"] += float(actual_split[2]) * rate
        row["unmapped_qty"] += int(detail.get("cost_unmapped_qty") or 0)
        row["unmapped_revenue_cny"] += (
            float(detail.get("cost_unmapped_revenue_pln") or 0) * rate
        )
        row["site_count"] += int(partner["site_count"] or 0)
        partner_profit_ratio = (
            0.25
            if partner["partner_profit_ratio"] is None
            else float(partner["partner_profit_ratio"])
        )
        our_profit_ratio = (
            0.25
            if partner["our_profit_ratio"] is None
            else float(partner["our_profit_ratio"])
        )
        margin_ratio_total = partner_profit_ratio + our_profit_ratio
        row["configured_margin_share_rates"].append(
            our_profit_ratio / margin_ratio_total
            if margin_ratio_total > 0
            else 0.5
        )

    for row in basis.values():
        net = row["reconciliation_net_cny"]
        row["contract_effective_rate"] = (
            row["contract_company_revenue_cny"] / net
            if abs(net) > 0.01
            else None
        )
        row["actual_cost_rate"] = (
            row["actual_cost_cny"] / net if abs(net) > 0.01 else None
        )
        configured_rates = row.pop("configured_margin_share_rates")
        row["our_margin_share_rate"] = (
            sum(configured_rates) / len(configured_rates)
            if configured_rates
            else None
        )
        row["cost_complete"] = row["unmapped_qty"] == 0
        for key in (
            "reconciliation_net_cny",
            "actual_cost_cny",
            "actual_margin_cny",
            "contract_company_revenue_cny",
            "actual_company_revenue_cny",
            "unmapped_revenue_cny",
        ):
            row[key] = round(row[key], 2)

    return basis, gaps


def _latest_actual_cost_forecast_rates(conn, year_month):
    """Return the latest confirmed actual-cost economics before a month.

    Actual monthly results remain fail-closed until that month's reconciliation
    snapshot exists. Forecasts may safely reuse the latest earlier snapshot,
    while retaining the configured gross-margin split instead of reverse
    engineering a percentage from rounded statement amounts.
    """
    statement_columns = _table_columns(conn, "reconciliation_statements")
    required = {
        "partner_id",
        "period_year",
        "period_month",
        "total_net_pln",
        "actual_cost_pln_snapshot",
        "exchange_rate_cny",
        "is_manual",
    }
    if not required.issubset(statement_columns):
        return {}

    partner_sites_relation = _partner_sites_relation(conn)
    partners = conn.execute(
        f"""
        SELECT p.id, p.currency, p.partner_profit_ratio, p.our_profit_ratio,
               GROUP_CONCAT(DISTINCT s.country) AS countries
        FROM partners p
        LEFT JOIN {partner_sites_relation} ps ON ps.partner_id = p.id
        LEFT JOIN sites s ON s.id = ps.site_id
        GROUP BY p.id
        """
    ).fetchall()
    result = {}
    for partner in partners:
        countries = sorted(
            {
                value.strip().upper()
                for value in (partner["countries"] or "").split(",")
                if value.strip()
            }
        )
        if len(countries) != 1:
            continue
        snapshot = conn.execute(
            """
            SELECT period_year, period_month, total_net_pln,
                   actual_cost_pln_snapshot, exchange_rate_cny
            FROM reconciliation_statements
            WHERE partner_id = ?
              AND printf('%04d-%02d', period_year, period_month) < ?
              AND actual_cost_pln_snapshot IS NOT NULL
              AND total_net_pln > 0
              AND COALESCE(is_manual, 0) = 0
            ORDER BY period_year DESC, period_month DESC, id DESC
            LIMIT 1
            """,
            (partner["id"], year_month),
        ).fetchone()
        if not snapshot:
            continue
        rate = snapshot["exchange_rate_cny"]
        if not rate:
            rate = _rate_to_cny(
                conn,
                partner["currency"],
                f'{snapshot["period_year"]}-{snapshot["period_month"]:02d}',
            )
        if not rate:
            continue

        net_cny = float(snapshot["total_net_pln"] or 0) * float(rate)
        cost_cny = float(snapshot["actual_cost_pln_snapshot"] or 0) * float(rate)
        partner_ratio = (
            0.25
            if partner["partner_profit_ratio"] is None
            else float(partner["partner_profit_ratio"])
        )
        our_ratio = (
            0.25
            if partner["our_profit_ratio"] is None
            else float(partner["our_profit_ratio"])
        )
        ratio_total = partner_ratio + our_ratio
        margin_share = our_ratio / ratio_total if ratio_total > 0 else 0.5
        company_revenue_cny = (net_cny - cost_cny) * margin_share

        country = countries[0]
        row = result.setdefault(
            country,
            {
                "net_cny": 0.0,
                "cost_cny": 0.0,
                "company_revenue_cny": 0.0,
                "margin_share_rates": [],
                "source_months": [],
            },
        )
        row["net_cny"] += net_cny
        row["cost_cny"] += cost_cny
        row["company_revenue_cny"] += company_revenue_cny
        row["margin_share_rates"].append(margin_share)
        row["source_months"].append(
            f'{snapshot["period_year"]}-{snapshot["period_month"]:02d}'
        )

    for row in result.values():
        net_cny = row["net_cny"]
        rates = row.pop("margin_share_rates")
        row["actual_cost_rate"] = (
            row["cost_cny"] / net_cny if abs(net_cny) > 0.01 else None
        )
        row["effective_company_revenue_rate"] = (
            row["company_revenue_cny"] / net_cny
            if abs(net_cny) > 0.01
            else None
        )
        row["our_margin_share_rate"] = (
            sum(rates) / len(rates) if rates else None
        )
        row["source_months"] = sorted(set(row["source_months"]))
    return result


def _partner_contract_rates(conn):
    """Return lightweight configured company shares for historical trends."""
    partner_sites_relation = _partner_sites_relation(conn)
    rows = conn.execute(
        f"""
        SELECT s.country, p.our_profit_ratio
        FROM partners p
        JOIN {partner_sites_relation} ps ON ps.partner_id = p.id
        JOIN sites s ON s.id = ps.site_id
        WHERE s.country IS NOT NULL AND s.country != ''
        GROUP BY p.id, s.country
        """
    ).fetchall()
    grouped = defaultdict(list)
    for row in rows:
        grouped[row["country"].upper()].append(
            0.25
            if row["our_profit_ratio"] is None
            else float(row["our_profit_ratio"])
        )

    result = {}
    for country, rates in grouped.items():
        if rates and max(rates) - min(rates) <= 1e-9:
            result[country] = rates[0]
    return result


def _expense_rows(conn, year_month):
    return [
        dict(row)
        for row in conn.execute(
            """
            SELECT id, year_month, scenario, category, name, amount_cny,
                   is_recurring, notes, created_by, created_at, updated_at
            FROM company_profit_expenses
            WHERE year_month = ?
            ORDER BY scenario, category, id
            """,
            (year_month,),
        ).fetchall()
    ]


def _forecast_recurring_rows(conn, year_month, direct_forecast_rows):
    direct_keys = {
        (row["category"].strip().casefold(), row["name"].strip().casefold())
        for row in direct_forecast_rows
    }
    rows = conn.execute(
        """
        SELECT id, year_month, category, name, amount_cny, notes
        FROM company_profit_expenses
        WHERE scenario = 'actual'
          AND is_recurring = 1
          AND year_month < ?
        ORDER BY year_month DESC, id DESC
        """,
        (year_month,),
    ).fetchall()
    seen = set()
    result = []
    for row in rows:
        key = (row["category"].strip().casefold(), row["name"].strip().casefold())
        if key in seen or key in direct_keys:
            continue
        seen.add(key)
        item = dict(row)
        item["source_month"] = item.pop("year_month")
        result.append(item)
    return result


def _historical_boards(compute_sales_board, year_month, count=3):
    result = []
    for offset in range(1, count + 1):
        month = _month_shift(year_month, -offset)
        result.append((month, compute_sales_board(month)))
    return result


def _weighted(values):
    weights = (0.5, 0.3, 0.2)
    pairs = [
        (float(value or 0), weights[index])
        for index, value in enumerate(values[: len(weights)])
    ]
    weight_total = sum(weight for _, weight in pairs)
    return sum(value * weight for value, weight in pairs) / weight_total if pairs else 0


def _latest_targets(conn, year_month):
    rows = conn.execute(
        """
        SELECT st.manager, st.base_salary, st.commission_rate
        FROM sales_targets st
        JOIN (
            SELECT manager, MAX(year_month) AS year_month
            FROM sales_targets
            WHERE year_month <= ?
            GROUP BY manager
        ) latest
          ON latest.manager = st.manager
         AND latest.year_month = st.year_month
        """,
        (year_month,),
    ).fetchall()
    return {row["manager"]: dict(row) for row in rows}


def _forecast_payroll(
    conn,
    year_month,
    board,
    forecast_manager_net,
    historical_boards,
):
    targets = _latest_targets(conn, year_month)
    latest_board_by_manager = {
        row["manager"]: row for row in board.get("board_data", [])
    }
    for _, history in historical_boards:
        for row in history.get("board_data", []):
            latest_board_by_manager.setdefault(row["manager"], row)

    forecast_commission_base = {}
    base_total = 0.0
    commission_total = 0.0
    missing_targets = []
    # Full-month payroll must include every configured salaried person, even
    # when their forecast sales are zero.  Commission-only collaborators stay
    # at zero base because that is what their sales target records.
    payroll_managers = sorted(set(targets) | set(forecast_manager_net))
    for manager in payroll_managers:
        forecast_net = float(forecast_manager_net.get(manager, 0) or 0)
        target = targets.get(manager)
        if not target:
            if forecast_net > 0:
                missing_targets.append(manager)
            continue
        base_total += float(target["base_salary"] or 0)
        history_row = latest_board_by_manager.get(manager, {})
        history_net = float(history_row.get("month_net_cny") or 0)
        history_base = float(history_row.get("commission_base_cny") or 0)
        eligible_ratio = history_base / history_net if history_net > 0 else 1.0
        eligible_ratio = min(max(eligible_ratio, 0), 1.25)
        predicted_base = float(forecast_net or 0) * eligible_ratio
        forecast_commission_base[manager] = predicted_base
        commission_total += predicted_base * float(target["commission_rate"] or 0)

    leader_bonus = 0.0
    groups = conn.execute(
        "SELECT id, leader_manager, bonus_rate FROM sales_groups"
    ).fetchall()
    for group in groups:
        members = [
            row["manager"]
            for row in conn.execute(
                "SELECT manager FROM sales_group_members WHERE group_id = ?",
                (group["id"],),
            ).fetchall()
        ]
        bonus_base = sum(
            forecast_commission_base.get(manager, 0)
            for manager in members
            if manager != group["leader_manager"]
        )
        leader_bonus += bonus_base * float(group["bonus_rate"] or 0)

    return {
        "amount_cny": round(base_total + commission_total + leader_bonus, 2),
        "base_salary_cny": round(base_total, 2),
        "commission_cny": round(commission_total, 2),
        "leader_bonus_cny": round(leader_bonus, 2),
        "missing_salary_rules": sorted(missing_targets),
        "method": "全额底薪 + 预测提成 + 预测带团奖金（不预设扣薪）",
    }


def _trend_sales(conn, revenue_status_cond, months):
    start_month, end_month = months[0], months[-1]
    result = {
        month: {
            "gmv_cny": 0.0,
            "net_by_country": defaultdict(float),
            "missing_rates": set(),
        }
        for month in months
    }

    gmv_rows = conn.execute(
        """
        SELECT strftime('%Y-%m', o.date_created) AS year_month,
               o.currency,
               SUM(o.total) AS amount
        FROM orders o
        WHERE strftime('%Y-%m', o.date_created) BETWEEN ? AND ?
          AND o.status NOT IN ('checkout-draft', 'trash')
        GROUP BY year_month, o.currency
        """,
        (start_month, end_month),
    ).fetchall()
    for row in gmv_rows:
        rate = _rate_to_cny(conn, row["currency"], row["year_month"])
        if rate is None:
            result[row["year_month"]]["missing_rates"].add(
                (row["currency"] or "N/A").upper()
            )
        else:
            result[row["year_month"]]["gmv_cny"] += float(row["amount"] or 0) * rate

    net_rows = conn.execute(
        f"""
        SELECT strftime('%Y-%m', o.date_created) AS year_month,
               COALESCE(s.country, '未分配') AS country,
               o.currency,
               SUM(o.total - o.shipping_total) AS amount
        FROM orders o
        LEFT JOIN sites s ON s.url = o.source
        WHERE strftime('%Y-%m', o.date_created) BETWEEN ? AND ?
          AND {revenue_status_cond('o')}
        GROUP BY year_month, country, o.currency
        """,
        (start_month, end_month),
    ).fetchall()
    for row in net_rows:
        rate = _rate_to_cny(conn, row["currency"], row["year_month"])
        if rate is None:
            result[row["year_month"]]["missing_rates"].add(
                (row["currency"] or "N/A").upper()
            )
        else:
            result[row["year_month"]]["net_by_country"][row["country"]] += (
                float(row["amount"] or 0) * rate
            )

    loss_rows = conn.execute(
        """
        SELECT strftime('%Y-%m', o.date_created) AS year_month,
               COALESCE(s.country, '未分配') AS country,
               o.currency,
               SUM(
                   COALESCE(o.shipping_loss_amount, 0)
                   + CASE WHEN COALESCE(o.is_problem_return, 0) = 1
                          THEN COALESCE(o.product_loss_amount, 0)
                          ELSE 0 END
               ) AS amount
        FROM orders o
        LEFT JOIN sites s ON s.url = o.source
        WHERE strftime('%Y-%m', o.date_created) BETWEEN ? AND ?
          AND (
              COALESCE(o.is_undelivered, 0) = 1
              OR COALESCE(o.is_problem_return, 0) = 1
          )
        GROUP BY year_month, country, o.currency
        """,
        (start_month, end_month),
    ).fetchall()
    for row in loss_rows:
        rate = _rate_to_cny(conn, row["currency"], row["year_month"])
        if rate is None:
            result[row["year_month"]]["missing_rates"].add(
                (row["currency"] or "N/A").upper()
            )
        else:
            result[row["year_month"]]["net_by_country"][row["country"]] -= (
                float(row["amount"] or 0) * rate
            )
    return result


def build_company_profit_summary(
    get_db_connection,
    compute_sales_board,
    revenue_status_cond,
    year_month,
    today=None,
    trend_count=14,
    partner_recon_detail=None,
    statement_split=None,
    prefer_reconciled_snapshots=False,
):
    year_month = _valid_month(year_month)
    today = today or date.today()
    current_month = today.strftime("%Y-%m")
    # One company-profit request can need the same month's sales board for the
    # selected summary, three-month forecast history, and trend payroll. Keep a
    # request-local cache so each month is calculated at most once.
    board_cache = {}

    def board_for_month(month):
        if month not in board_cache:
            board_cache[month] = compute_sales_board(month)
        return board_cache[month]

    board = board_for_month(year_month)
    historical_boards = _historical_boards(board_for_month, year_month)

    conn = get_db_connection()
    try:
        settings = _month_settings(conn, year_month)
        rules = _market_rules(conn, year_month)
        calculation_mode = settings["calculation_mode"]
        fallback_country_basis = (
            board.get("team_totals", {}).get("country_profit", {})
            if prefer_reconciled_snapshots
            else None
        )
        partner_basis, partner_basis_gaps = _partner_reconciliation_basis(
            conn,
            year_month,
            partner_recon_detail,
            statement_split,
            fallback_country_basis=fallback_country_basis,
            fallback_has_actual_cost=False,
        )
        historical_actual_cost_rates = _latest_actual_cost_forecast_rates(
            conn, year_month
        )
        partner_contract_rates = _partner_contract_rates(conn)
        expenses = _expense_rows(conn, year_month)
        actual_expenses = [
            row for row in expenses if row["scenario"] == "actual"
        ]
        forecast_expenses = [
            row for row in expenses if row["scenario"] == "forecast"
        ]
        recurring_forecast = _forecast_recurring_rows(
            conn, year_month, forecast_expenses
        )

        gmv_cny, gmv_by_country, missing_rates = _month_gmv(conn, year_month)
        board_countries = board.get("team_totals", {}).get("country_profit", {})
        saved_countries = set(rules)
        site_countries = {
            row["country"]
            for row in conn.execute(
                """
                SELECT DISTINCT country
                FROM sites
                WHERE country IS NOT NULL AND country != ''
                """
            ).fetchall()
        }
        countries = sorted(
            set(board_countries)
            | saved_countries
            | site_countries
            | set(partner_basis)
            | set(historical_actual_cost_rates)
        )

        selected_year, selected_month_number = (
            int(part) for part in year_month.split("-")
        )
        days_in_month = monthrange(selected_year, selected_month_number)[1]
        elapsed_days = (
            min(today.day, days_in_month)
            if year_month == current_month
            else days_in_month
        )

        history_country_values = defaultdict(list)
        history_manager_values = defaultdict(list)
        for _, history in historical_boards:
            history_country = history.get("team_totals", {}).get(
                "country_profit", {}
            )
            for country in countries:
                history_country_values[country].append(
                    float(history_country.get(country, {}).get("net_cny") or 0)
                )
            history_managers = {
                row["manager"]: float(row.get("month_net_cny") or 0)
                for row in history.get("board_data", [])
            }
            all_managers = {
                row["manager"] for row in board.get("board_data", [])
            } | set(history_managers)
            for manager in all_managers:
                history_manager_values[manager].append(
                    history_managers.get(manager, 0)
                )

        actual_company_revenue = 0.0
        forecast_company_revenue = 0.0
        forecast_net_sales_cny = 0.0
        team_net_sales_cny = float(
            board.get("team_totals", {}).get("month_net_cny") or 0
        )
        settlement_net_sales_cny = 0.0
        country_rows = []
        missing_actual_basis_countries = []
        missing_forecast_basis_countries = []
        incomplete_actual_cost_countries = []
        forecast_manager_net = {}

        actual_manager_net = {
            row["manager"]: float(row.get("month_net_cny") or 0)
            for row in board.get("board_data", [])
        }
        for manager, actual_net in actual_manager_net.items():
            if year_month == current_month and elapsed_days:
                forecast_manager_net[manager] = (
                    actual_net / elapsed_days * days_in_month
                )
            elif year_month > current_month:
                forecast_manager_net[manager] = _weighted(
                    history_manager_values.get(manager, [])
                )
            else:
                forecast_manager_net[manager] = actual_net

        for country in countries:
            board_country = board_countries.get(country, {})
            net_cny = float(board_country.get("net_cny") or 0)
            rule = rules.get(country, {})
            reconciliation = partner_basis.get(country)
            historical_actual_cost = historical_actual_cost_rates.get(country)

            manual_forecast = rule.get("forecast_net_sales_cny")
            if manual_forecast is not None:
                forecast_net = float(manual_forecast)
                forecast_source = "手工输入"
            elif year_month == current_month and elapsed_days:
                forecast_net = net_cny / elapsed_days * days_in_month
                forecast_source = "本月至今日均外推"
            elif year_month > current_month:
                forecast_net = _weighted(history_country_values.get(country, []))
                forecast_source = "近3个月加权"
            else:
                forecast_net = net_cny
                forecast_source = "历史实际"

            share_locked = bool(reconciliation)
            actual_product_cost_cny = None
            actual_cost_rate = None
            unmapped_cost_qty = 0
            unmapped_cost_revenue_cny = 0.0
            reconciliation_net_cny = None
            partner_names = []
            reconciliation_sources = []
            statement_statuses = []
            statement_updated_at = []

            if reconciliation:
                reconciliation_net_cny = reconciliation[
                    "reconciliation_net_cny"
                ]
                actual_product_cost_cny = (
                    reconciliation["actual_cost_cny"]
                    if reconciliation["actual_cost_available"]
                    else None
                )
                actual_cost_rate = (
                    reconciliation["actual_cost_rate"]
                    if reconciliation["actual_cost_available"]
                    else None
                )
                unmapped_cost_qty = reconciliation["unmapped_qty"]
                unmapped_cost_revenue_cny = reconciliation[
                    "unmapped_revenue_cny"
                ]
                partner_names = reconciliation["partner_names"]
                reconciliation_sources = reconciliation["basis_sources"]
                statement_statuses = reconciliation["statement_statuses"]
                statement_updated_at = reconciliation[
                    "statement_updated_at"
                ]

                if calculation_mode == "actual_cost":
                    share_rate = reconciliation["our_margin_share_rate"]
                    if not reconciliation["actual_cost_available"]:
                        share_source = "缺少合伙人对账单实际成本快照"
                        actual_revenue = None
                        historical_effective_rate = (
                            historical_actual_cost.get(
                                "effective_company_revenue_rate"
                            )
                            if historical_actual_cost
                            else None
                        )
                        if historical_effective_rate is not None:
                            forecast_revenue = (
                                forecast_net
                                * float(historical_effective_rate)
                            )
                            forecast_revenue_source = (
                                "沿用最近已确认月份实际成本率"
                            )
                        elif abs(forecast_net) <= 0.01:
                            forecast_revenue = 0.0
                            forecast_revenue_source = (
                                "本月无销售，无需成本外推"
                            )
                        else:
                            forecast_revenue = None
                            forecast_revenue_source = (
                                "缺少历史实际成本预测基础"
                            )
                    else:
                        share_source = "合伙人对账 · 实际成本"
                        actual_revenue = reconciliation[
                            "actual_company_revenue_cny"
                        ]
                    if reconciliation["actual_cost_available"]:
                        if not reconciliation["cost_complete"]:
                            incomplete_actual_cost_countries.append(
                                {
                                    "country": country,
                                    "qty": unmapped_cost_qty,
                                }
                            )
                        if (
                            actual_cost_rate is not None
                            and share_rate is not None
                        ):
                            forecast_revenue = (
                                forecast_net
                                * (1 - float(actual_cost_rate))
                                * float(share_rate)
                            )
                            forecast_revenue_source = (
                                "按本月实际成本率与剩余毛利分成外推"
                            )
                        elif abs(forecast_net) <= 0.01:
                            forecast_revenue = 0.0
                            forecast_revenue_source = (
                                "本月无销售，无需成本外推"
                            )
                        else:
                            forecast_revenue = None
                            forecast_revenue_source = (
                                "缺少实际成本预测基础"
                            )
                else:
                    share_rate = reconciliation["contract_effective_rate"]
                    share_source = "合伙人对账 · 约定比例"
                    actual_revenue = reconciliation[
                        "contract_company_revenue_cny"
                    ]
                    forecast_revenue = (
                        forecast_net * float(share_rate)
                        if share_rate is not None
                        else (0.0 if abs(forecast_net) <= 0.01 else None)
                    )
                    forecast_revenue_source = (
                        "按对账有效分润比例外推"
                        if share_rate is not None
                        else "本月无销售，无需比例外推"
                    )
            elif calculation_mode == "actual_cost":
                historical_effective_rate = (
                    historical_actual_cost.get(
                        "effective_company_revenue_rate"
                    )
                    if historical_actual_cost
                    else None
                )
                share_rate = (
                    historical_actual_cost.get("our_margin_share_rate")
                    if historical_actual_cost
                    else None
                )
                share_source = (
                    "最近已确认月份实际成本口径"
                    if historical_actual_cost
                    else "缺少合伙人实际成本"
                )
                actual_revenue = 0.0 if abs(net_cny) <= 0.01 else None
                if historical_effective_rate is not None:
                    forecast_revenue = (
                        forecast_net * float(historical_effective_rate)
                    )
                    forecast_revenue_source = (
                        "沿用最近已确认月份实际成本率"
                    )
                else:
                    forecast_revenue = (
                        0.0 if abs(forecast_net) <= 0.01 else None
                    )
                    forecast_revenue_source = "缺少历史实际成本预测基础"
                if actual_revenue is None:
                    missing_actual_basis_countries.append(country)
                if forecast_revenue is None:
                    missing_forecast_basis_countries.append(country)
            else:
                share_rate = rule.get("share_rate")
                share_source = "月度设置"
                if share_rate is None and country in DEFAULT_SHARE_RATES:
                    share_rate = DEFAULT_SHARE_RATES[country]
                    share_source = "业务约定默认值"
                if share_rate is None and abs(net_cny) > 0.01:
                    missing_actual_basis_countries.append(country)
                if share_rate is None and abs(forecast_net) > 0.01:
                    missing_forecast_basis_countries.append(country)
                actual_revenue = (
                    net_cny * float(share_rate)
                    if share_rate is not None
                    else None
                )
                forecast_revenue = (
                    forecast_net * float(share_rate)
                    if share_rate is not None
                    else None
                )
                forecast_revenue_source = (
                    f"按{share_source}外推"
                    if share_rate is not None
                    else "缺少分润比例"
                )

            if reconciliation and actual_revenue is None:
                missing_actual_basis_countries.append(country)
            if reconciliation and forecast_revenue is None:
                missing_forecast_basis_countries.append(country)

            if actual_revenue is not None:
                actual_company_revenue += actual_revenue
            if forecast_revenue is not None:
                forecast_company_revenue += forecast_revenue
            forecast_net_sales_cny += forecast_net
            settlement_net_cny = (
                reconciliation_net_cny
                if reconciliation_net_cny is not None
                else net_cny
            )
            settlement_net_sales_cny += settlement_net_cny

            country_rows.append(
                {
                    "country": country,
                    "gmv_cny": round(gmv_by_country.get(country, 0), 2),
                    "net_sales_cny": round(net_cny, 2),
                    "settlement_net_sales_cny": round(
                        settlement_net_cny, 2
                    ),
                    "reconciliation_net_cny": reconciliation_net_cny,
                    "actual_product_cost_cny": actual_product_cost_cny,
                    "actual_cost_rate": actual_cost_rate,
                    "unmapped_cost_qty": unmapped_cost_qty,
                    "unmapped_cost_revenue_cny": round(
                        unmapped_cost_revenue_cny, 2
                    ),
                    "partner_names": partner_names,
                    "reconciliation_sources": reconciliation_sources,
                    "statement_statuses": statement_statuses,
                    "statement_updated_at": statement_updated_at,
                    "share_rate": share_rate,
                    "share_locked": share_locked,
                    "share_source": share_source,
                    "company_revenue_cny": (
                        round(actual_revenue, 2)
                        if actual_revenue is not None
                        else None
                    ),
                    "forecast_net_sales_cny": round(forecast_net, 2),
                    "forecast_source": forecast_source,
                    "forecast_revenue_source": forecast_revenue_source,
                    "forecast_company_revenue_cny": (
                        round(forecast_revenue, 2)
                        if forecast_revenue is not None
                        else None
                    ),
                    "manual_forecast_net_sales_cny": manual_forecast,
                    "notes": rule.get("notes", ""),
                }
            )

        system_payroll = float(
            board.get("team_totals", {}).get("total_income") or 0
        )
        if settings["payroll_actual_override"] is not None:
            actual_payroll = float(settings["payroll_actual_override"])
            payroll_source = "实际已发（手工确认）"
        else:
            actual_payroll = system_payroll
            payroll_source = "销售看板动态暂估"

        forecast_payroll_detail = _forecast_payroll(
            conn,
            year_month,
            board,
            forecast_manager_net,
            historical_boards,
        )
        if settings["payroll_forecast_override"] is not None:
            forecast_payroll = float(settings["payroll_forecast_override"])
            forecast_payroll_source = "手工输入"
        else:
            forecast_payroll = forecast_payroll_detail["amount_cny"]
            forecast_payroll_source = "系统规则预测"

        actual_other_expenses = sum(
            float(row["amount_cny"] or 0) for row in actual_expenses
        )
        inherited_forecast_expenses = sum(
            float(row["amount_cny"] or 0) for row in forecast_expenses
        ) + sum(float(row["amount_cny"] or 0) for row in recurring_forecast)
        if year_month <= current_month and actual_expenses:
            forecast_other_expenses = actual_other_expenses
            forecast_other_expenses_source = "当月实际支出"
        else:
            forecast_other_expenses = inherited_forecast_expenses
            forecast_other_expenses_source = "当月预测与周期性继承支出"

        actual_profit = (
            actual_company_revenue - actual_payroll - actual_other_expenses
        )
        forecast_profit = (
            forecast_company_revenue
            - forecast_payroll
            - forecast_other_expenses
        )
        default_company_revenue_rate = (
            forecast_company_revenue / forecast_net_sales_cny
            if abs(forecast_net_sales_cny) > 0.01
            else 0.0
        )
        scenario = _forecast_scenario(
            conn,
            year_month,
            forecast_net_sales_cny,
            default_company_revenue_rate,
            forecast_payroll,
            forecast_other_expenses,
        )
        if not scenario["is_saved"]:
            scenario["fixed_cost_source"] = forecast_other_expenses_source
        is_historical_month = year_month < current_month
        is_current_month = year_month == current_month
        has_settlement_sales = abs(settlement_net_sales_cny) > 0.01
        ladder_base_net_sales = (
            settlement_net_sales_cny
            if is_historical_month and has_settlement_sales
            else forecast_net_sales_cny
        )
        revenue_ladder = _revenue_ladder(
            year_month,
            ladder_base_net_sales,
            scenario["company_revenue_rate"],
            forecast_net_sales_cny,
            forecast_payroll_detail,
            settings["payroll_forecast_override"],
            forecast_other_expenses,
            forecast_other_expenses_source,
            (
                actual_payroll
                if is_historical_month
                else None
            ),
            (
                payroll_source
                if is_historical_month
                else None
            ),
            (
                settlement_net_sales_cny
                if is_current_month and has_settlement_sales
                else None
            ),
            (
                actual_payroll
                if is_current_month and has_settlement_sales
                else None
            ),
            (
                payroll_source
                if is_current_month and has_settlement_sales
                else None
            ),
        )
        revenue_ladder["base_source"] = (
            "本月合伙人结算净销售"
            if is_historical_month and has_settlement_sales
            else (
                "本月月底预测净销售"
                if is_current_month
                else "系统预测净销售"
            )
        )
        actual_complete = bool(
            not missing_actual_basis_countries
            and not incomplete_actual_cost_countries
            and not partner_basis_gaps
            and settings["payroll_actual_override"] is not None
            and settings["actual_expenses_complete"]
            and not missing_rates
        )
        forecast_complete = bool(
            not missing_forecast_basis_countries
            and not incomplete_actual_cost_countries
            and not partner_basis_gaps
            and (
                settings["payroll_forecast_override"] is not None
                or not forecast_payroll_detail["missing_salary_rules"]
            )
            and not missing_rates
        )

        trend_months = _month_sequence(year_month, trend_count)
        trend_sales = _trend_sales(conn, revenue_status_cond, trend_months)
        expense_totals = {
            row["year_month"]: float(row["total"] or 0)
            for row in conn.execute(
                """
                SELECT year_month, SUM(amount_cny) AS total
                FROM company_profit_expenses
                WHERE scenario = 'actual'
                  AND year_month BETWEEN ? AND ?
                GROUP BY year_month
                """,
                (trend_months[0], trend_months[-1]),
            ).fetchall()
        }
        month_settings = {
            row["year_month"]: dict(row)
            for row in conn.execute(
                """
                SELECT *
                FROM company_profit_month_settings
                WHERE year_month BETWEEN ? AND ?
                """,
                (trend_months[0], trend_months[-1]),
            ).fetchall()
        }
        rule_rows = conn.execute(
            """
            SELECT year_month, country, share_rate
            FROM company_profit_market_rules
            WHERE year_month BETWEEN ? AND ?
            """,
            (trend_months[0], trend_months[-1]),
        ).fetchall()
        trend_rules = defaultdict(dict)
        for row in rule_rows:
            trend_rules[row["year_month"]][row["country"]] = row["share_rate"]

        target_months = {
            row["year_month"]
            for row in conn.execute(
                """
                SELECT DISTINCT year_month
                FROM sales_targets
                WHERE year_month BETWEEN ? AND ?
                """,
                (trend_months[0], trend_months[-1]),
            ).fetchall()
        }
        trend = []
        for month in trend_months:
            sales = trend_sales[month]
            month_setting = month_settings.get(month, {})
            month_mode = month_setting.get("calculation_mode") or "percentage"
            if month_mode not in CALCULATION_MODES:
                month_mode = "percentage"
            if month == year_month:
                month_partner_basis = partner_basis
                month_partner_gaps = partner_basis_gaps
            else:
                month_partner_basis, month_partner_gaps = (
                    _partner_reconciliation_basis(
                        conn,
                        month,
                        partner_recon_detail,
                        statement_split,
                        fallback_country_basis={
                            country: {"net_cny": net}
                            for country, net in sales[
                                "net_by_country"
                            ].items()
                        },
                        fallback_has_actual_cost=False,
                    )
                )
            company_revenue = 0.0
            settlement_net = 0.0
            missing_basis = list(month_partner_gaps)
            for country in sorted(
                set(sales["net_by_country"]) | set(month_partner_basis)
            ):
                net = sales["net_by_country"].get(country, 0.0)
                reconciliation = month_partner_basis.get(country)
                if reconciliation:
                    settlement_net += reconciliation[
                        "reconciliation_net_cny"
                    ]
                    if month_mode == "actual_cost":
                        if reconciliation["actual_cost_available"]:
                            company_revenue += reconciliation[
                                "actual_company_revenue_cny"
                            ]
                        else:
                            missing_basis.append(
                                f"{country}缺少实际成本快照"
                            )
                        if (
                            reconciliation["actual_cost_available"]
                            and not reconciliation["cost_complete"]
                        ):
                            missing_basis.append(
                                f"{country}实际成本未完整匹配"
                            )
                    else:
                        company_revenue += reconciliation[
                            "contract_company_revenue_cny"
                        ]
                    continue
                settlement_net += net

                if month_mode == "actual_cost":
                    if abs(net) > 0.01:
                        missing_basis.append(f"{country}缺少实际成本")
                    continue

                share = trend_rules[month].get(country)
                if share is None:
                    share = partner_contract_rates.get(country)
                if share is None:
                    share = DEFAULT_SHARE_RATES.get(country)
                if share is None and abs(net) > 0.01:
                    missing_basis.append(f"{country}缺少分润比例")
                elif share is not None:
                    company_revenue += net * float(share)

            payroll = month_setting.get("payroll_actual_override")
            payroll_source_for_trend = "manual" if payroll is not None else None
            if payroll is None and month in target_months:
                payroll = float(
                    board_for_month(month)
                    .get("team_totals", {})
                    .get("total_income")
                    or 0
                )
                payroll_source_for_trend = "system"
            other_expense = expense_totals.get(month, 0.0)
            provisional_profit = (
                company_revenue - float(payroll) - other_expense
                if payroll is not None and not missing_basis
                else None
            )
            complete = bool(
                provisional_profit is not None
                and month_setting.get("payroll_actual_override") is not None
                and month_setting.get("actual_expenses_complete")
                and not sales["missing_rates"]
            )
            trend.append(
                {
                    "month": month,
                    "gmv_cny": round(sales["gmv_cny"], 2),
                    "net_sales_cny": round(settlement_net, 2),
                    "team_net_sales_cny": round(
                        sum(sales["net_by_country"].values()), 2
                    ),
                    "company_revenue_cny": round(company_revenue, 2),
                    "calculation_mode": month_mode,
                    "profit_cny": (
                        round(provisional_profit, 2)
                        if provisional_profit is not None
                        else None
                    ),
                    "complete": complete,
                    "payroll_source": payroll_source_for_trend,
                }
            )

        gaps = []
        gaps.extend(partner_basis_gaps)
        if missing_actual_basis_countries:
            basis_name = (
                "合伙人对账单实际成本快照"
                if calculation_mode == "actual_cost"
                else "市场分润比例"
            )
            gaps.append(
                f"实际数据缺少{basis_name}："
                + "、".join(missing_actual_basis_countries)
            )
        forecast_only_missing = sorted(
            set(missing_forecast_basis_countries)
            - set(missing_actual_basis_countries)
        )
        if forecast_only_missing:
            basis_name = (
                "合伙人对账单实际成本快照"
                if calculation_mode == "actual_cost"
                else "市场分润比例"
            )
            gaps.append(
                f"预测数据缺少{basis_name}："
                + "、".join(forecast_only_missing)
            )
        for item in incomplete_actual_cost_countries:
            gaps.append(
                f'{item["country"]}实际成本未完整匹配：'
                f'{item["qty"]}件商品缺少有效进价'
            )
        if settings["payroll_actual_override"] is None:
            gaps.append("未录入实际已发工资，当前采用销售看板动态暂估")
        if not settings["actual_expenses_complete"]:
            gaps.append("尚未确认本月其他支出已录完")
        if missing_rates:
            gaps.append("缺少汇率：" + "、".join(missing_rates))
        if (
            settings["payroll_forecast_override"] is None
            and forecast_payroll_detail["missing_salary_rules"]
        ):
            gaps.append(
                "缺少工资预测规则："
                + "、".join(forecast_payroll_detail["missing_salary_rules"])
            )

        return {
            "year_month": year_month,
            "calculation_mode": calculation_mode,
            "calculation_mode_label": (
                "按实际成本"
                if calculation_mode == "actual_cost"
                else "按分润比例"
            ),
            "is_current_month": year_month == current_month,
            "can_close_actual": year_month <= current_month,
            "gmv_cny": gmv_cny,
            "net_sales_cny": round(settlement_net_sales_cny, 2),
            "team_net_sales_cny": round(team_net_sales_cny, 2),
            "actual": {
                "company_revenue_cny": round(actual_company_revenue, 2),
                "payroll_cny": round(actual_payroll, 2),
                "system_payroll_cny": round(system_payroll, 2),
                "payroll_source": payroll_source,
                "other_expenses_cny": round(actual_other_expenses, 2),
                "profit_cny": round(actual_profit, 2),
                "margin_pct": (
                    round(actual_profit / actual_company_revenue * 100, 1)
                    if actual_company_revenue
                    else None
                ),
                "complete": actual_complete,
            },
            "forecast": {
                "net_sales_cny": round(forecast_net_sales_cny, 2),
                "company_revenue_cny": round(forecast_company_revenue, 2),
                "payroll_cny": round(forecast_payroll, 2),
                "payroll_source": forecast_payroll_source,
                "payroll_detail": forecast_payroll_detail,
                "other_expenses_cny": round(forecast_other_expenses, 2),
                "other_expenses_source": forecast_other_expenses_source,
                "profit_cny": round(forecast_profit, 2),
                "margin_pct": (
                    round(forecast_profit / forecast_company_revenue * 100, 1)
                    if forecast_company_revenue
                    else None
                ),
                "complete": forecast_complete,
                "method_note": (
                    "当月按已过天数外推；未来月份按近3个月 50%/30%/20% 加权。"
                    "手工输入时优先使用手工值。"
                    + (
                        "实际成本市场优先按本月成本率外推；当月对账快照尚未生成时，沿用最近已确认月份的实际成本率。"
                        if calculation_mode == "actual_cost"
                        else "合伙人市场按对账有效分润比例外推。"
                    )
                ),
            },
            "countries": country_rows,
            "expenses": expenses,
            "recurring_forecast_expenses": recurring_forecast,
            "settings": settings,
            "scenario": scenario,
            "revenue_ladder": revenue_ladder,
            "data_gaps": gaps,
            "trend": trend,
            "definitions": {
                "gmv": "非草稿/非回收站订单总额，含失败和取消订单，用于观察市场需求规模。",
                "net_sales": (
                    "合伙人结算净销售优先引用冻结对账单；"
                    "团队业绩净销售来自销售看板，用于提成和绩效。"
                ),
                "company_revenue": (
                    "按实际成本：合伙人先收回真实产品成本，剩余毛利按对账设置分配；"
                    "公司可得收入直接引用合伙人对账结果。"
                    if calculation_mode == "actual_cost"
                    else "按分润比例：公司可得收入直接引用合伙人对账的约定比例结果；"
                    "未绑定合伙人的市场才使用本页月度分润设置。"
                ),
                "actual_profit": "公司可得收入－实际/暂估工资－已录入实际支出。",
                "forecast_profit": "预测公司可得收入－全额底薪及预测提成－预测/周期性支出。",
                "scenario_profit": "阶梯公司收入＝目标结算净销售额×公司收入率；工资按全额底薪加随销售额变化的提成与带团奖金测算；日常运营汇总本月预测及周期性支出。",
            },
        }
    finally:
        conn.close()
