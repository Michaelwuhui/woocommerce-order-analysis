#!/usr/bin/env python3
"""Create, load, and verify PostgreSQL from an authoritative SQLite snapshot.

The source is always opened read-only.  Production reset requires an explicit
flag; normal development targets a database whose name ends in _test.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
import re
import sqlite3
import sys
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any

import psycopg


ROOT = Path(__file__).resolve().parents[2]
MIGRATION_DIR = ROOT / "migrations" / "postgresql"
OWNER_ROLE = "woo_analysis_owner"
APP_ROLE = "woo_analysis_app"

JSON_TEXT_COLUMNS = {
    ("orders", name)
    for name in (
        "billing", "shipping", "meta_data", "line_items", "tax_lines",
        "shipping_lines", "fee_lines", "coupon_lines",
    )
} | {
    ("orders_archive", name)
    for name in (
        "billing", "shipping", "meta_data", "line_items", "tax_lines",
        "shipping_lines", "fee_lines", "coupon_lines",
    )
}

# The live audit proved these contain composite site/order identifiers despite
# their historical INTEGER declarations.
TEXT_TYPE_OVERRIDES = {
    ("order_notes", "order_id"), ("shipping_logs", "order_id"),
}

BOOLEAN_EXACT = {
    "enabled", "emailed", "customer_note", "added_by_user",
    "prices_include_tax", "delivery_confirmed", "copy_to_fallback",
    "routed_to_master", "actual_expenses_complete",
}
BOOLEAN_PREFIXES = ("is_", "has_", "can_")
BOOLEAN_SUFFIXES = ("_enabled", "_active", "_required", "_confirmed", "_complete")


def quote_ident(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def safe_name(prefix: str, *parts: str) -> str:
    raw = "_".join((prefix,) + tuple(parts))
    cleaned = re.sub(r"[^a-zA-Z0-9_]+", "_", raw).lower()
    if len(cleaned) <= 55:
        return cleaned
    digest = hashlib.sha256(cleaned.encode()).hexdigest()[:7]
    return cleaned[:55] + "_" + digest


@dataclass(frozen=True)
class Column:
    cid: int
    name: str
    declared_type: str
    notnull: bool
    default: str | None
    pk_position: int
    hidden: int


@dataclass
class Table:
    name: str
    sql: str
    columns: list[Column]
    autoincrement: bool
    boolean_columns: set[str]


class Catalog:
    def __init__(self, db: sqlite3.Connection):
        self.db = db
        self.db.row_factory = sqlite3.Row
        self.tables: list[Table] = []
        for row in db.execute(
            "SELECT name,sql FROM sqlite_schema "
            "WHERE type='table' AND name!='sqlite_sequence' ORDER BY name"
        ):
            columns = [
                Column(
                    int(item["cid"]),
                    str(item["name"]),
                    str(item["type"] or ""),
                    bool(item["notnull"]),
                    item["dflt_value"],
                    int(item["pk"]),
                    int(item["hidden"]),
                )
                for item in db.execute(
                    "PRAGMA table_xinfo(" + quote_ident(row["name"]) + ")"
                )
                if int(item["hidden"]) == 0
            ]
            booleans = {
                column.name
                for column in columns
                if self._is_boolean(row["name"], column)
            }
            self.tables.append(
                Table(
                    str(row["name"]),
                    str(row["sql"] or ""),
                    columns,
                    "AUTOINCREMENT" in str(row["sql"] or "").upper(),
                    booleans,
                )
            )
        self.by_name = {table.name: table for table in self.tables}

    def _is_boolean(self, table: str, column: Column) -> bool:
        declared = column.declared_type.upper()
        if declared not in {"INTEGER", "INT"}:
            return False
        name = column.name.lower()
        semantic = (
            name in BOOLEAN_EXACT
            or name.startswith(BOOLEAN_PREFIXES)
            or name.endswith(BOOLEAN_SUFFIXES)
        )
        if not semantic or name.endswith("_by") or name.endswith("_id"):
            return False
        ident = quote_ident(column.name)
        table_ident = quote_ident(table)
        values = [
            row[0]
            for row in self.db.execute(
                "SELECT DISTINCT " + ident + " FROM " + table_ident
                + " WHERE " + ident + " IS NOT NULL LIMIT 4"
            )
        ]
        return all(value in (0, 1) for value in values)

    def primary_key(self, table: Table) -> tuple[str, ...]:
        return tuple(
            column.name
            for column in sorted(table.columns, key=lambda item: item.pk_position or 9999)
            if column.pk_position
        )

    def unique_column_sets(self, table: Table) -> list[tuple[str, ...]]:
        result: list[tuple[str, ...]] = []
        for index in self.db.execute(
            "PRAGMA index_list(" + quote_ident(table.name) + ")"
        ):
            if not int(index["unique"]) or str(index["origin"]) == "pk":
                continue
            columns = [
                str(item["name"])
                for item in self.db.execute(
                    "PRAGMA index_xinfo(" + quote_ident(index["name"]) + ")"
                )
                if int(item["key"]) and int(item["cid"]) >= 0
            ]
            key_columns = [
                item
                for item in self.db.execute(
                    "PRAGMA index_xinfo(" + quote_ident(index["name"]) + ")"
                )
                if int(item["key"])
            ]
            if columns and len(columns) == len(key_columns):
                value = tuple(columns)
                if value not in result:
                    result.append(value)
        return result

    def explicit_indexes(self, table: Table):
        for index in self.db.execute(
            "PRAGMA index_list(" + quote_ident(table.name) + ")"
        ):
            if str(index["origin"]) != "c":
                continue
            row = self.db.execute(
                "SELECT sql FROM sqlite_schema WHERE type='index' AND name=?",
                (index["name"],),
            ).fetchone()
            if row and row[0]:
                yield str(index["name"]), str(row[0])

    def foreign_keys(self, table: Table):
        grouped: dict[int, list[sqlite3.Row]] = defaultdict(list)
        for row in self.db.execute(
            "PRAGMA foreign_key_list(" + quote_ident(table.name) + ")"
        ):
            grouped[int(row["id"])].append(row)
        for fk_id, rows in sorted(grouped.items()):
            rows.sort(key=lambda item: int(item["seq"]))
            yield fk_id, rows


def sqlite_connection(path: Path) -> sqlite3.Connection:
    uri = "file:" + str(path.resolve()) + "?mode=ro"
    connection = sqlite3.connect(uri, uri=True)
    connection.execute("PRAGMA query_only=ON")
    connection.row_factory = sqlite3.Row
    return connection


def declared_pg_type(table: Table, column: Column) -> str:
    if (table.name, column.name) in TEXT_TYPE_OVERRIDES:
        return "text"
    if column.name in table.boolean_columns:
        return "boolean"
    declared = column.declared_type.upper().strip()
    if (table.name, column.name) in JSON_TEXT_COLUMNS or column.name.lower().endswith("_json"):
        # Kept as text for compatibility with json.loads and exact source bytes.
        return "text"
    if "INT" in declared:
        return "bigint"
    if any(token in declared for token in ("REAL", "FLOA", "DOUB", "NUMERIC", "DECIMAL")):
        return "numeric"
    if "BLOB" in declared:
        return "bytea"
    if declared in {"DATETIME", "TIMESTAMP"}:
        return "timestamp without time zone"
    return "text"


def translated_default(table: Table, column: Column, pg_type: str) -> str | None:
    value = column.default
    if value is None:
        return None
    stripped = str(value).strip()
    if column.name in table.boolean_columns:
        if stripped.strip("()'\"").lower() in {"1", "true"}:
            return "true"
        if stripped.strip("()'\"").lower() in {"0", "false"}:
            return "false"
    if stripped.upper() == "CURRENT_TIMESTAMP" or stripped.lower() in {
        "datetime('now')", "(datetime('now'))"
    }:
        if pg_type.startswith("timestamp"):
            return "CURRENT_TIMESTAMP"
        return "to_char(timezone('UTC', CURRENT_TIMESTAMP), 'YYYY-MM-DD HH24:MI:SS')"
    if stripped.upper() in {"CURRENT_DATE", "CURRENT_TIME"}:
        return stripped.upper()
    if re.fullmatch(r"[-+]?[0-9]+(?:\.[0-9]+)?", stripped):
        return stripped
    if stripped.startswith("'") and stripped.endswith("'"):
        return stripped
    return stripped


def extract_checks(create_sql: str) -> list[str]:
    checks: list[str] = []
    upper = create_sql.upper()
    position = 0
    while True:
        found = upper.find("CHECK", position)
        if found < 0:
            break
        cursor = found + 5
        while cursor < len(create_sql) and create_sql[cursor].isspace():
            cursor += 1
        if cursor >= len(create_sql) or create_sql[cursor] != "(":
            position = cursor
            continue
        depth = 1
        start = cursor + 1
        cursor += 1
        quote = None
        while cursor < len(create_sql) and depth:
            char = create_sql[cursor]
            if quote:
                if char == quote and cursor + 1 < len(create_sql) and create_sql[cursor + 1] == quote:
                    cursor += 2
                    continue
                if char == quote:
                    quote = None
            elif char in "'\"":
                quote = char
            elif char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    checks.append(create_sql[start:cursor].strip())
                    break
            cursor += 1
        position = cursor + 1
    return checks


def translate_check(table: Table, expression: str) -> str:
    result = expression
    for name in table.boolean_columns:
        token = re.escape(name)
        result = re.sub(
            r"\b" + token + r"\b\s+IN\s*\(\s*0\s*,\s*1\s*\)",
            quote_ident(name) + " IN (false,true)",
            result,
            flags=re.IGNORECASE,
        )
        result = re.sub(
            r"\b" + token + r"\b\s*=\s*1\b",
            quote_ident(name) + " = true",
            result,
            flags=re.IGNORECASE,
        )
        result = re.sub(
            r"\b" + token + r"\b\s*=\s*0\b",
            quote_ident(name) + " = false",
            result,
            flags=re.IGNORECASE,
        )
    return result


def table_ddl(catalog: Catalog, table: Table) -> str:
    pk = catalog.primary_key(table)
    pieces: list[str] = []
    for column in table.columns:
        pg_type = declared_pg_type(table, column)
        identity = (
            table.autoincrement
            and len(pk) == 1
            and pk[0] == column.name
            and pg_type == "bigint"
        )
        definition = quote_ident(column.name) + " " + pg_type
        if identity:
            definition += " GENERATED BY DEFAULT AS IDENTITY"
        default = translated_default(table, column, pg_type)
        if default is not None and not identity:
            definition += " DEFAULT " + default
        if column.notnull or column.name in pk:
            definition += " NOT NULL"
        pieces.append(definition)
    if pk:
        pieces.append(
            "CONSTRAINT " + quote_ident(safe_name("pk", table.name))
            + " PRIMARY KEY (" + ", ".join(quote_ident(name) for name in pk) + ")"
        )
    for position, columns in enumerate(catalog.unique_column_sets(table), start=1):
        if columns == pk:
            continue
        pieces.append(
            "CONSTRAINT " + quote_ident(safe_name("uq", table.name, str(position)))
            + " UNIQUE (" + ", ".join(quote_ident(name) for name in columns) + ")"
        )
    for position, expression in enumerate(extract_checks(table.sql), start=1):
        pieces.append(
            "CONSTRAINT " + quote_ident(safe_name("ck", table.name, str(position)))
            + " CHECK (" + translate_check(table, expression) + ")"
        )
    return (
        "CREATE TABLE " + quote_ident(table.name) + " (\n    "
        + ",\n    ".join(pieces) + "\n)"
    )


def convert_value(table: Table, column: Column, value: Any):
    if value is None:
        return None
    if column.name in table.boolean_columns:
        return bool(value)
    pg_type = declared_pg_type(table, column)
    if pg_type == "numeric":
        if isinstance(value, float):
            if not math.isfinite(value):
                raise ValueError("non-finite REAL in " + table.name + "." + column.name)
            return Decimal(repr(value))
        return Decimal(str(value))
    if pg_type.startswith("timestamp"):
        if isinstance(value, dt.datetime):
            return value.replace(tzinfo=None)
        raw = str(value).strip()
        if not raw:
            return None
        return dt.datetime.fromisoformat(raw.replace("Z", "+00:00")).replace(tzinfo=None)
    return value


def run_sql_file(pg, path: Path):
    sql = path.read_text(encoding="utf-8")
    with pg.cursor() as cursor:
        cursor.execute(sql, prepare=False)
    pg.commit()


def reset_schema(pg, database: str, allow_production_reset: bool):
    if not database.endswith("_test") and not allow_production_reset:
        raise RuntimeError(
            "Refusing schema reset outside *_test without --allow-production-reset"
        )
    pg.execute("DROP SCHEMA IF EXISTS public CASCADE")
    pg.execute("CREATE SCHEMA public AUTHORIZATION " + quote_ident(OWNER_ROLE))
    pg.execute("REVOKE CREATE ON SCHEMA public FROM PUBLIC")
    pg.execute("GRANT USAGE ON SCHEMA public TO " + quote_ident(APP_ROLE))
    pg.commit()


def create_tables(pg, catalog: Catalog):
    pg.execute("SET ROLE " + quote_ident(OWNER_ROLE))
    pg.execute("SET search_path TO public")
    for table in catalog.tables:
        pg.execute(table_ddl(catalog, table))
    pg.commit()


def import_table(pg, source: sqlite3.Connection, table: Table) -> int:
    columns = table.columns
    column_sql = ", ".join(quote_ident(column.name) for column in columns)
    select_sql = "SELECT " + column_sql + " FROM " + quote_ident(table.name)
    copy_sql = (
        "COPY " + quote_ident(table.name) + " (" + column_sql + ") FROM STDIN"
    )
    count = 0
    with pg.cursor() as cursor:
        with cursor.copy(copy_sql) as copy:
            for row in source.execute(select_sql):
                copy.write_row(
                    [
                        convert_value(table, column, row[column.name])
                        for column in columns
                    ]
                )
                count += 1
    pg.commit()
    return count


def translate_index_sql(sql: str) -> str:
    result = sql.strip().rstrip(";")
    result = re.sub(
        r"^CREATE\s+(UNIQUE\s+)?INDEX\s+",
        lambda match: "CREATE " + (match.group(1) or "") + "INDEX IF NOT EXISTS ",
        result,
        flags=re.IGNORECASE,
    )
    result = re.sub(r"\bdatetime\s*\(", "sqlite_datetime(", result, flags=re.IGNORECASE)
    result = re.sub(r"\bdate\s*\(", "sqlite_date(", result, flags=re.IGNORECASE)
    result = re.sub(r"\bstrftime\s*\(", "sqlite_strftime(", result, flags=re.IGNORECASE)
    return result


def add_indexes_foreign_keys_views(pg, source: sqlite3.Connection, catalog: Catalog):
    pg.execute("SET ROLE " + quote_ident(OWNER_ROLE))
    for table in catalog.tables:
        for _name, sql in catalog.explicit_indexes(table):
            pg.execute(translate_index_sql(sql))
    for table in catalog.tables:
        for fk_id, rows in catalog.foreign_keys(table):
            parent = str(rows[0]["table"])
            from_columns = [str(row["from"]) for row in rows]
            to_columns = [str(row["to"]) for row in rows]
            clause = (
                "ALTER TABLE " + quote_ident(table.name)
                + " ADD CONSTRAINT " + quote_ident(safe_name("fk", table.name, str(fk_id)))
                + " FOREIGN KEY (" + ", ".join(quote_ident(name) for name in from_columns) + ")"
                + " REFERENCES " + quote_ident(parent)
                + " (" + ", ".join(quote_ident(name) for name in to_columns) + ")"
            )
            on_update = str(rows[0]["on_update"] or "NO ACTION").upper()
            on_delete = str(rows[0]["on_delete"] or "NO ACTION").upper()
            if on_update != "NO ACTION":
                clause += " ON UPDATE " + on_update
            if on_delete != "NO ACTION":
                clause += " ON DELETE " + on_delete
            if table.name == "reconciliation_audit_log" and parent == "reconciliation_statements":
                clause += " NOT VALID"
            pg.execute(clause)

    for row in source.execute(
        "SELECT name,sql FROM sqlite_schema WHERE type='view' ORDER BY name"
    ):
        view_sql = str(row["sql"]).strip().rstrip(";")
        view_sql = re.sub(
            r"^CREATE\s+VIEW\s+",
            "CREATE OR REPLACE VIEW ",
            view_sql,
            flags=re.IGNORECASE,
        )
        pg.execute(view_sql)

    pg.execute(
        """
        CREATE OR REPLACE VIEW sqlite_master AS
        SELECT CASE cls.relkind
                   WHEN 'r' THEN 'table'
                   WHEN 'p' THEN 'table'
                   WHEN 'v' THEN 'view'
                   WHEN 'm' THEN 'view'
                   WHEN 'i' THEN 'index'
                   ELSE cls.relkind::text
               END AS type,
               cls.relname::text AS name,
               COALESCE(parent.relname, cls.relname)::text AS tbl_name,
               0::bigint AS rootpage,
               NULL::text AS sql
        FROM pg_class cls
        JOIN pg_namespace ns ON ns.oid=cls.relnamespace
        LEFT JOIN pg_index idx ON idx.indexrelid=cls.oid
        LEFT JOIN pg_class parent ON parent.oid=idx.indrelid
        WHERE ns.nspname=current_schema()
          AND cls.relkind IN ('r','p','v','m','i')
        """
    )
    pg.commit()


def reset_sequences(pg, source: sqlite3.Connection, catalog: Catalog):
    source_sequences = {
        str(row["name"]): int(row["seq"])
        for row in source.execute("SELECT name,seq FROM sqlite_sequence")
    }
    pg.execute("SET ROLE " + quote_ident(OWNER_ROLE))
    for table in catalog.tables:
        pk = catalog.primary_key(table)
        if not (table.autoincrement and len(pk) == 1):
            continue
        sequence = pg.execute(
            "SELECT pg_get_serial_sequence(%s,%s)",
            (table.name, pk[0]),
        ).fetchone()[0]
        if not sequence:
            continue
        target_max = pg.execute(
            "SELECT max(" + quote_ident(pk[0]) + ") FROM " + quote_ident(table.name)
        ).fetchone()[0]
        source_value = source_sequences.get(table.name, 0)
        value = max(int(target_max or 0), source_value)
        if value > 0:
            pg.execute("SELECT setval(%s::regclass,%s,true)", (sequence, value))
        else:
            pg.execute("SELECT setval(%s::regclass,1,false)", (sequence,))
    pg.commit()


def canonical(value: Any):
    if value is None:
        return None
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, float):
        return format(Decimal(repr(value)), "f")
    if isinstance(value, (dt.date, dt.datetime)):
        return value.isoformat(sep=" ")
    if isinstance(value, bool):
        # SQLite stores booleans as integer 0/1; use the same canonical text
        # representation so row digests compare semantic values, not JSON types.
        return "1" if value else "0"
    if isinstance(value, memoryview):
        return bytes(value).hex()
    if isinstance(value, bytes):
        return value.hex()
    return str(value)


def digest_key_rows(rows) -> str:
    # Database collations may order the same text key set differently. Compare
    # a Python-sorted canonical set so the digest measures membership only.
    encoded = sorted(
        json.dumps(
            [canonical(value) for value in row],
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode()
        for row in rows
    )
    digest = hashlib.sha256()
    for item in encoded:
        digest.update(item)
        digest.update(b"\n")
    return digest.hexdigest()


def key_digest_sqlite(source, table: Table, pk: tuple[str, ...]) -> str:
    query = (
        "SELECT " + ",".join(quote_ident(name) for name in pk)
        + " FROM " + quote_ident(table.name)
    )
    return digest_key_rows(source.execute(query))


def key_digest_postgres(pg, table: Table, pk: tuple[str, ...]) -> str:
    query = (
        "SELECT " + ",".join(quote_ident(name) for name in pk)
        + " FROM " + quote_ident(table.name)
    )
    with pg.cursor() as cursor:
        cursor.execute(query)
        return digest_key_rows(cursor)


def rows_digest_sqlite(source, table: Table, columns: list[str] | None = None) -> str:
    names = columns or [column.name for column in table.columns]
    query = (
        "SELECT " + ",".join(quote_ident(name) for name in names)
        + " FROM " + quote_ident(table.name)
    )
    return digest_key_rows(source.execute(query))


def rows_digest_postgres(pg, table: Table, columns: list[str] | None = None) -> str:
    names = columns or [column.name for column in table.columns]
    query = (
        "SELECT " + ",".join(quote_ident(name) for name in names)
        + " FROM " + quote_ident(table.name)
    )
    with pg.cursor() as cursor:
        cursor.execute(query)
        return digest_key_rows(cursor)


def postgres_unique_sets(pg, table_name: str) -> set[tuple[str, ...]]:
    rows = pg.execute(
        """
        SELECT array_agg(att.attname ORDER BY ord.n)
        FROM pg_index idx
        JOIN pg_class cls ON cls.oid=idx.indrelid
        JOIN pg_namespace ns ON ns.oid=cls.relnamespace
        JOIN LATERAL unnest(idx.indkey) WITH ORDINALITY ord(attnum,n) ON TRUE
        JOIN pg_attribute att
          ON att.attrelid=cls.oid AND att.attnum=ord.attnum
        WHERE ns.nspname=current_schema() AND cls.relname=%s
          AND idx.indisunique AND ord.attnum>0
        GROUP BY idx.indexrelid
        """,
        (table_name,),
    ).fetchall()
    return {tuple(str(value) for value in row[0]) for row in rows}


def postgres_fk_orphan_count(pg, table: Table, rows: list[sqlite3.Row]) -> int:
    parent = str(rows[0]["table"])
    joins = []
    nonnull = []
    for row in rows:
        child_name = str(row["from"])
        parent_name = str(row["to"])
        joins.append(
            "child." + quote_ident(child_name)
            + "=parent." + quote_ident(parent_name)
        )
        nonnull.append("child." + quote_ident(child_name) + " IS NOT NULL")
    first_parent = quote_ident(str(rows[0]["to"]))
    query = (
        "SELECT count(*) FROM " + quote_ident(table.name) + " child "
        + "LEFT JOIN " + quote_ident(parent) + " parent ON "
        + " AND ".join(joins)
        + " WHERE " + " AND ".join(nonnull)
        + " AND parent." + first_parent + " IS NULL"
    )
    return int(pg.execute(query).fetchone()[0])


def verify(pg, source: sqlite3.Connection, catalog: Catalog) -> dict[str, Any]:
    source_fk_issues = source.execute("PRAGMA foreign_key_check").fetchall()
    source_fk_counts: dict[tuple[str, int], int] = defaultdict(int)
    for issue in source_fk_issues:
        source_fk_counts[(str(issue["table"]), int(issue["fkid"]))] += 1

    report: dict[str, Any] = {
        "integrity_check": source.execute("PRAGMA integrity_check").fetchone()[0],
        "tables": {},
        "key_digests": {},
        "row_digests": {},
        "column_digests": {},
        "real_columns": {},
        "boolean_columns": {},
        "currency_columns": {},
        "foreign_keys": {},
        "sequences": {},
        "catalog": {},
        "foreign_key_issues": len(source_fk_issues),
        "ok": True,
    }
    for table in catalog.tables:
        pk = catalog.primary_key(table)
        source_count = source.execute(
            "SELECT count(*) FROM " + quote_ident(table.name)
        ).fetchone()[0]
        target_count = pg.execute(
            "SELECT count(*) FROM " + quote_ident(table.name)
        ).fetchone()[0]
        match = int(source_count) == int(target_count)
        report["tables"][table.name] = {
            "sqlite": int(source_count),
            "postgres": int(target_count),
            "match": match,
        }
        report["ok"] = report["ok"] and match

        left_rows = rows_digest_sqlite(source, table)
        right_rows = rows_digest_postgres(pg, table)
        report["row_digests"][table.name] = {
            "sqlite": left_rows,
            "postgres": right_rows,
            "match": left_rows == right_rows,
        }
        report["ok"] = report["ok"] and left_rows == right_rows

        if left_rows != right_rows:
            column_results = {}
            for column in table.columns:
                digest_columns = list(pk)
                if column.name not in digest_columns:
                    digest_columns.append(column.name)
                left_column = rows_digest_sqlite(source, table, digest_columns)
                right_column = rows_digest_postgres(pg, table, digest_columns)
                if left_column != right_column:
                    column_results[column.name] = {
                        "sqlite": left_column,
                        "postgres": right_column,
                        "match": False,
                    }
            report["column_digests"][table.name] = column_results

        if pk:
            left = key_digest_sqlite(source, table, pk)
            right = key_digest_postgres(pg, table, pk)
            report["key_digests"][table.name] = {
                "sqlite": left,
                "postgres": right,
                "match": left == right,
            }
            report["ok"] = report["ok"] and left == right

        for column in table.columns:
            if column.name in table.boolean_columns:
                source_values = source.execute(
                    "SELECT count(*) FILTER (WHERE " + quote_ident(column.name) + " IS NULL),"
                    + "count(*) FILTER (WHERE " + quote_ident(column.name) + "=0),"
                    + "count(*) FILTER (WHERE " + quote_ident(column.name) + "=1)"
                    + " FROM " + quote_ident(table.name)
                ).fetchone()
                target_values = pg.execute(
                    "SELECT count(*) FILTER (WHERE " + quote_ident(column.name) + " IS NULL),"
                    + "count(*) FILTER (WHERE " + quote_ident(column.name) + "=false),"
                    + "count(*) FILTER (WHERE " + quote_ident(column.name) + "=true)"
                    + " FROM " + quote_ident(table.name)
                ).fetchone()
                left_bool = [int(value) for value in source_values]
                right_bool = [int(value) for value in target_values]
                bool_match = left_bool == right_bool
                report["boolean_columns"][table.name + "." + column.name] = {
                    "sqlite_null_false_true": left_bool,
                    "postgres_null_false_true": right_bool,
                    "match": bool_match,
                }
                report["ok"] = report["ok"] and bool_match

            if "currency" in column.name.lower():
                digest_columns = list(pk) + [column.name] if pk else [column.name]
                left_currency = rows_digest_sqlite(source, table, digest_columns)
                right_currency = rows_digest_postgres(pg, table, digest_columns)
                currency_match = left_currency == right_currency
                report["currency_columns"][table.name + "." + column.name] = {
                    "sqlite": left_currency,
                    "postgres": right_currency,
                    "match": currency_match,
                }
                report["ok"] = report["ok"] and currency_match

            if declared_pg_type(table, column) != "numeric":
                continue
            source_values = [
                Decimal(repr(row[0]))
                for row in source.execute(
                    "SELECT " + quote_ident(column.name)
                    + " FROM " + quote_ident(table.name)
                    + " WHERE " + quote_ident(column.name) + " IS NOT NULL"
                )
            ]
            source_sum = sum(source_values, Decimal(0))
            target_sum = pg.execute(
                "SELECT COALESCE(sum(" + quote_ident(column.name) + "),0)"
                + " FROM " + quote_ident(table.name)
            ).fetchone()[0]
            digest_columns = list(pk) + [column.name] if pk else [column.name]
            left_numeric = rows_digest_sqlite(source, table, digest_columns)
            right_numeric = rows_digest_postgres(pg, table, digest_columns)
            match = source_sum == Decimal(target_sum) and left_numeric == right_numeric
            report["real_columns"][table.name + "." + column.name] = {
                "sqlite_sum": format(source_sum, "f"),
                "postgres_sum": format(Decimal(target_sum), "f"),
                "sqlite_digest": left_numeric,
                "postgres_digest": right_numeric,
                "match": match,
            }
            report["ok"] = report["ok"] and match

    expected_indexes = {
        name
        for table in catalog.tables
        for name, _sql in catalog.explicit_indexes(table)
    }
    actual_indexes = {
        str(row[0])
        for row in pg.execute(
            "SELECT indexname FROM pg_indexes WHERE schemaname=current_schema()"
        )
    }
    missing_indexes = sorted(expected_indexes - actual_indexes)
    report["catalog"]["indexes"] = {
        "expected": len(expected_indexes),
        "present": len(expected_indexes & actual_indexes),
        "missing": missing_indexes,
        "match": not missing_indexes,
    }
    report["ok"] = report["ok"] and not missing_indexes

    unique_mismatches = {}
    for table in catalog.tables:
        expected_unique = set(catalog.unique_column_sets(table))
        pk = catalog.primary_key(table)
        if pk:
            expected_unique.add(pk)
        actual_unique = postgres_unique_sets(pg, table.name)
        missing = sorted(expected_unique - actual_unique)
        if missing:
            unique_mismatches[table.name] = [list(item) for item in missing]
    report["catalog"]["unique_constraints"] = {
        "tables_checked": len(catalog.tables),
        "missing": unique_mismatches,
        "match": not unique_mismatches,
    }
    report["ok"] = report["ok"] and not unique_mismatches

    target_fk_constraints = {
        str(row[0]): bool(row[1])
        for row in pg.execute(
            """
            SELECT con.conname,con.convalidated
            FROM pg_constraint con
            JOIN pg_namespace ns ON ns.oid=con.connamespace
            WHERE ns.nspname=current_schema() AND con.contype='f'
            """
        )
    }
    target_orphans = 0
    expected_fk_count = 0
    for table in catalog.tables:
        for fk_id, fk_rows in catalog.foreign_keys(table):
            expected_fk_count += 1
            parent = str(fk_rows[0]["table"])
            constraint = safe_name("fk", table.name, str(fk_id))
            source_orphans = source_fk_counts.get((table.name, fk_id), 0)
            postgres_orphans = postgres_fk_orphan_count(pg, table, fk_rows)
            target_orphans += postgres_orphans
            expected_validated = not (
                table.name == "reconciliation_audit_log"
                and parent == "reconciliation_statements"
            )
            present = constraint in target_fk_constraints
            validated = target_fk_constraints.get(constraint)
            fk_match = (
                present
                and validated == expected_validated
                and source_orphans == postgres_orphans
            )
            report["foreign_keys"][constraint] = {
                "table": table.name,
                "parent": parent,
                "sqlite_orphans": source_orphans,
                "postgres_orphans": postgres_orphans,
                "present": present,
                "validated": validated,
                "expected_validated": expected_validated,
                "match": fk_match,
            }
            report["ok"] = report["ok"] and fk_match
    report["postgres_foreign_key_issues"] = target_orphans
    report["catalog"]["foreign_keys"] = {
        "expected": expected_fk_count,
        "present": sum(
            1 for name in report["foreign_keys"]
            if name in target_fk_constraints
        ),
        "match": all(item["match"] for item in report["foreign_keys"].values()),
    }

    source_views = {
        str(row[0])
        for row in source.execute(
            "SELECT name FROM sqlite_schema WHERE type='view'"
        )
    }
    target_views = {
        str(row[0])
        for row in pg.execute(
            """
            SELECT cls.relname FROM pg_class cls
            JOIN pg_namespace ns ON ns.oid=cls.relnamespace
            WHERE ns.nspname=current_schema() AND cls.relkind IN ('v','m')
            """
        )
    }
    missing_views = sorted(source_views - target_views)
    report["catalog"]["views"] = {
        "expected": sorted(source_views),
        "missing": missing_views,
        "match": not missing_views,
    }
    report["ok"] = report["ok"] and not missing_views

    source_sequences = {
        str(row["name"]): int(row["seq"])
        for row in source.execute("SELECT name,seq FROM sqlite_sequence")
    }
    for table in catalog.tables:
        pk = catalog.primary_key(table)
        if not (table.autoincrement and len(pk) == 1):
            continue
        sequence = pg.execute(
            "SELECT pg_get_serial_sequence(%s,%s)", (table.name, pk[0])
        ).fetchone()[0]
        target_max = int(pg.execute(
            "SELECT COALESCE(max(" + quote_ident(pk[0]) + "),0) FROM "
            + quote_ident(table.name)
        ).fetchone()[0])
        source_value = source_sequences.get(table.name, 0)
        expected_next = max(target_max, source_value) + 1
        schema_name, sequence_name = str(sequence).split(".", 1)
        sequence_name = sequence_name.strip('"')
        state = pg.execute(
            "SELECT last_value,is_called FROM "
            + quote_ident(schema_name.strip('"')) + "." + quote_ident(sequence_name)
        ).fetchone()
        actual_next = int(state[0]) + (1 if bool(state[1]) else 0)
        sequence_match = actual_next == expected_next
        report["sequences"][table.name] = {
            "sequence": str(sequence),
            "source_sequence": source_value,
            "table_max": target_max,
            "expected_next": expected_next,
            "actual_next": actual_next,
            "match": sequence_match,
        }
        report["ok"] = report["ok"] and sequence_match
    return report


def write_report(path: Path, report: dict[str, Any]):
    flags = os.O_WRONLY | os.O_CREAT | os.O_TRUNC
    fd = os.open(path, flags, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")


def grant_runtime(pg):
    pg.execute("SET ROLE " + quote_ident(OWNER_ROLE))
    pg.execute(
        "GRANT SELECT,INSERT,UPDATE,DELETE ON ALL TABLES IN SCHEMA public TO "
        + quote_ident(APP_ROLE)
    )
    pg.execute(
        "GRANT USAGE,SELECT,UPDATE ON ALL SEQUENCES IN SCHEMA public TO "
        + quote_ident(APP_ROLE)
    )
    pg.execute("RESET ROLE")
    pg.commit()


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sqlite", required=True, type=Path)
    parser.add_argument("--database", required=True)
    parser.add_argument("--dsn")
    parser.add_argument("--reset", action="store_true")
    parser.add_argument("--allow-production-reset", action="store_true")
    parser.add_argument("--verify-only", action="store_true")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args(argv)

    if not args.sqlite.is_file():
        parser.error("SQLite snapshot does not exist")
    source = sqlite_connection(args.sqlite)
    catalog = Catalog(source)
    dsn = args.dsn or "dbname=" + args.database
    pg = psycopg.connect(dsn, autocommit=False)

    try:
        if args.verify_only:
            if args.reset:
                parser.error("--verify-only cannot be combined with --reset")
            report = verify(pg, source, catalog)
            report["database"] = args.database
            report["source"] = str(args.sqlite)
            if args.report:
                write_report(args.report, report)
            print(json.dumps({"ok": report["ok"], "tables": len(catalog.tables)}))
            return 0 if report["ok"] else 2
        if args.reset:
            reset_schema(pg, args.database, args.allow_production_reset)
        run_sql_file(pg, MIGRATION_DIR / "001_compatibility.sql")
        create_tables(pg, catalog)
        imported = {}
        for table in catalog.tables:
            imported[table.name] = import_table(pg, source, table)
            print("imported", table.name, imported[table.name], flush=True)
        add_indexes_foreign_keys_views(pg, source, catalog)
        reset_sequences(pg, source, catalog)
        run_sql_file(pg, MIGRATION_DIR / "002_sync_pipeline.sql")
        run_sql_file(pg, MIGRATION_DIR / "003_sync_post_commit.sql")
        grant_runtime(pg)
        report = verify(pg, source, catalog)
        report["database"] = args.database
        report["source"] = str(args.sqlite)
        report["imported"] = imported
        if args.report:
            write_report(args.report, report)
        print(json.dumps({"ok": report["ok"], "tables": len(catalog.tables)}))
        return 0 if report["ok"] else 2
    finally:
        pg.close()
        source.close()


if __name__ == "__main__":
    raise SystemExit(main())
