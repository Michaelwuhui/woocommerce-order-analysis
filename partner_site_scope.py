"""Country-inheriting site scope for partner reconciliation."""


EFFECTIVE_PARTNER_SITES = "effective_partner_sites"


def init_partner_site_scope(conn):
    """Create country grants, exclusions, and the live effective-scope view."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS partner_country_permissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            partner_id INTEGER NOT NULL,
            country TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(partner_id, country),
            FOREIGN KEY (partner_id) REFERENCES partners(id) ON DELETE CASCADE
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS partner_site_exclusions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            partner_id INTEGER NOT NULL,
            site_id INTEGER NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(partner_id, site_id),
            FOREIGN KEY (partner_id) REFERENCES partners(id) ON DELETE CASCADE
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_partner_country_scope "
        "ON partner_country_permissions(partner_id, country)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_partner_site_exclusions "
        "ON partner_site_exclusions(partner_id, site_id)"
    )
    conn.execute(f"DROP VIEW IF EXISTS {EFFECTIVE_PARTNER_SITES}")
    conn.execute(
        f"""
        CREATE VIEW {EFFECTIVE_PARTNER_SITES} AS
        SELECT ps.partner_id, ps.site_id
        FROM partner_sites ps
        WHERE NOT EXISTS (
            SELECT 1
            FROM partner_site_exclusions pse
            WHERE pse.partner_id = ps.partner_id
              AND pse.site_id = ps.site_id
        )
        UNION
        SELECT pcp.partner_id, s.id AS site_id
        FROM partner_country_permissions pcp
        JOIN sites s
          ON UPPER(COALESCE(s.country, '')) = UPPER(pcp.country)
        WHERE COALESCE(s.country, '') != ''
          AND NOT EXISTS (
              SELECT 1
              FROM partner_site_exclusions pse
              WHERE pse.partner_id = pcp.partner_id
                AND pse.site_id = s.id
          )
        """
    )


def get_partner_site_scope(conn, partner_id):
    """Return explicit, inherited, excluded, and effective partner site IDs."""
    explicit = {
        row["site_id"]
        for row in conn.execute(
            "SELECT site_id FROM partner_sites WHERE partner_id = ?",
            (partner_id,),
        ).fetchall()
    }
    countries = {
        (row["country"] or "").upper()
        for row in conn.execute(
            """
            SELECT country
            FROM partner_country_permissions
            WHERE partner_id = ?
            """,
            (partner_id,),
        ).fetchall()
        if row["country"]
    }
    exclusions = {
        row["site_id"]
        for row in conn.execute(
            """
            SELECT site_id
            FROM partner_site_exclusions
            WHERE partner_id = ?
            """,
            (partner_id,),
        ).fetchall()
    }
    effective = {
        row["site_id"]
        for row in conn.execute(
            f"""
            SELECT site_id
            FROM {EFFECTIVE_PARTNER_SITES}
            WHERE partner_id = ?
            """,
            (partner_id,),
        ).fetchall()
    }
    return {
        "explicit_site_ids": explicit,
        "granted_countries": countries,
        "excluded_site_ids": exclusions,
        "effective_site_ids": effective,
    }


def replace_partner_site_scope(
    conn,
    partner_id,
    site_ids=None,
    country_grants=None,
    site_exclusions=None,
):
    """Replace a partner's scope atomically on the caller's transaction.

    Explicit sites are retained only outside granted countries. Exclusions are
    retained only inside granted countries. A site cannot be effectively bound
    to two partners because that would double-count reconciliation orders.
    """
    partner = conn.execute(
        "SELECT id FROM partners WHERE id = ?",
        (partner_id,),
    ).fetchone()
    if not partner:
        raise ValueError("合伙人不存在")

    site_rows = conn.execute(
        "SELECT id, url, country FROM sites"
    ).fetchall()
    sites = {int(row["id"]): row for row in site_rows}
    available_countries = {
        (row["country"] or "").upper()
        for row in site_rows
        if row["country"]
    }

    def normalize_ids(values, label):
        result = set()
        for value in values or []:
            try:
                site_id = int(value)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"{label}包含无效站点") from exc
            if site_id not in sites:
                raise ValueError(f"{label}包含不存在的站点：{site_id}")
            result.add(site_id)
        return result

    explicit = normalize_ids(site_ids, "单站点绑定")
    exclusions = normalize_ids(site_exclusions, "排除站点")
    countries = {
        str(country or "").strip().upper()
        for country in (country_grants or [])
        if str(country or "").strip()
    }
    invalid_countries = countries - available_countries
    if invalid_countries:
        raise ValueError(
            "国家不存在：" + "、".join(sorted(invalid_countries))
        )

    inherited_ids = {
        site_id
        for site_id, site in sites.items()
        if (site["country"] or "").upper() in countries
    }
    explicit = {
        site_id
        for site_id in explicit
        if (sites[site_id]["country"] or "").upper() not in countries
    }
    exclusions &= inherited_ids
    candidate_ids = explicit | (inherited_ids - exclusions)

    conflicts = []
    if candidate_ids:
        placeholders = ",".join("?" for _ in candidate_ids)
        conflicts = conn.execute(
            f"""
            SELECT eps.site_id, p.id AS partner_id, p.name AS partner_name
            FROM {EFFECTIVE_PARTNER_SITES} eps
            JOIN partners p ON p.id = eps.partner_id
            WHERE eps.partner_id != ?
              AND eps.site_id IN ({placeholders})
            ORDER BY eps.site_id, p.id
            """,
            [partner_id, *sorted(candidate_ids)],
        ).fetchall()
    if conflicts:
        labels = []
        for row in conflicts:
            site = sites[int(row["site_id"])]
            labels.append(
                f'{site["url"]}（已属于{row["partner_name"]}）'
            )
        raise ValueError(
            "以下站点不能重复计入多个合伙人：" + "、".join(labels)
        )

    conn.execute(
        "DELETE FROM partner_sites WHERE partner_id = ?",
        (partner_id,),
    )
    conn.execute(
        "DELETE FROM partner_country_permissions WHERE partner_id = ?",
        (partner_id,),
    )
    conn.execute(
        "DELETE FROM partner_site_exclusions WHERE partner_id = ?",
        (partner_id,),
    )
    for site_id in sorted(explicit):
        conn.execute(
            """
            INSERT INTO partner_sites (partner_id, site_id)
            VALUES (?, ?)
            """,
            (partner_id, site_id),
        )
    for country in sorted(countries):
        conn.execute(
            """
            INSERT INTO partner_country_permissions (partner_id, country)
            VALUES (?, ?)
            """,
            (partner_id, country),
        )
    for site_id in sorted(exclusions):
        conn.execute(
            """
            INSERT INTO partner_site_exclusions (partner_id, site_id)
            VALUES (?, ?)
            """,
            (partner_id, site_id),
        )

    return {
        "explicit_site_ids": sorted(explicit),
        "granted_countries": sorted(countries),
        "excluded_site_ids": sorted(exclusions),
        "effective_site_ids": sorted(candidate_ids),
    }
