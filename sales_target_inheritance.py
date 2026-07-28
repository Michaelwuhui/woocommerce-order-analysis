"""Sales-target loading with month-safe salary-rule inheritance."""


def load_sales_targets_for_month(conn, year_month):
    """Return target rows keyed by manager for one month.

    Exact rows keep every saved monthly field.  Managers without an exact row
    inherit only the latest prior base salary and commission rate; monthly and
    weekly sales targets remain empty for the new month.
    """
    exact_rows = conn.execute(
        """
        SELECT *
        FROM sales_targets
        WHERE year_month = ?
        ORDER BY manager
        """,
        (year_month,),
    ).fetchall()
    targets = {}
    for row in exact_rows:
        item = dict(row)
        item["salary_inherited_from_month"] = None
        targets[item["manager"]] = item

    prior_rows = conn.execute(
        """
        SELECT st.*
        FROM sales_targets st
        JOIN (
            SELECT manager, MAX(year_month) AS year_month
            FROM sales_targets
            WHERE year_month < ?
            GROUP BY manager
        ) latest
          ON latest.manager = st.manager
         AND latest.year_month = st.year_month
        ORDER BY st.manager
        """,
        (year_month,),
    ).fetchall()
    for row in prior_rows:
        if row["manager"] in targets:
            continue
        item = dict(row)
        source_month = item["year_month"]
        item.update(
            {
                "id": None,
                "year_month": year_month,
                "monthly_target": 0,
                "weekly_targets": "{}",
                "notes": "",
                "created_at": None,
                "updated_at": None,
                "salary_inherited_from_month": source_month,
            }
        )
        targets[item["manager"]] = item

    return targets
