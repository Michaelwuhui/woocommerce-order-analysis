"""Create a read-only company-profit snapshot for the offline Excel exporter.

This command never runs calculations against the live database connection.
It takes a consistent SQLite backup through a read-only source connection,
boots the existing calculation engine against that temporary copy, and writes
one private JSON snapshot outside the application tree.
"""

from argparse import ArgumentParser
from contextlib import contextmanager
from datetime import datetime
import importlib
import json
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
from zoneinfo import ZoneInfo

from company_profit import (
    MONTH_RE,
    build_company_profit_summary,
    init_company_profit_tables,
)


APP_ROOT = Path(__file__).resolve().parent
DEFAULT_DATABASE = APP_ROOT / "woocommerce_orders.db"


def _is_relative_to(path, parent):
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def _validate_month(value):
    value = (value or "").strip()
    if not MONTH_RE.fullmatch(value):
        raise ValueError("月份必须使用 YYYY-MM 格式")
    return value


def _validate_output_path(value):
    output_path = Path(value).expanduser().resolve()
    if _is_relative_to(output_path, APP_ROOT):
        raise ValueError(
            "安全限制：离线财务快照不能写入网站应用目录"
        )
    if output_path.suffix.lower() != ".json":
        raise ValueError("快照输出文件必须使用 .json 扩展名")
    return output_path


def _sqlite_readonly_backup(source_path, destination_path):
    source_uri = f"{source_path.resolve().as_uri()}?mode=ro"
    source = sqlite3.connect(source_uri, uri=True)
    destination = sqlite3.connect(destination_path)
    try:
        source.backup(destination)
    finally:
        destination.close()
        source.close()


@contextmanager
def _working_directory(path):
    original = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(original)


def build_snapshot(month, database_path):
    month = _validate_month(month)
    database_path = Path(database_path).expanduser().resolve()
    if not database_path.is_file():
        raise FileNotFoundError(f"数据库不存在：{database_path}")

    with tempfile.TemporaryDirectory(
        prefix="company-profit-offline-",
        ignore_cleanup_errors=True,
    ) as tempdir:
        temp_root = Path(tempdir)
        snapshot_db = temp_root / "woocommerce_orders.db"
        _sqlite_readonly_backup(database_path, snapshot_db)

        with _working_directory(temp_root):
            sys.path.insert(0, str(APP_ROOT))
            try:
                app_module = importlib.import_module("app")
                with app_module.app.app_context():
                    init_company_profit_tables(app_module.get_db_connection)
                    summary = build_company_profit_summary(
                        app_module.get_db_connection,
                        app_module._compute_sales_board_data,
                        app_module._revenue_status_cond,
                        month,
                        partner_recon_detail=app_module._calc_partner_recon_detail,
                        statement_split=app_module._compute_statement_split,
                        prefer_reconciled_snapshots=True,
                    )
            finally:
                if sys.path and sys.path[0] == str(APP_ROOT):
                    sys.path.pop(0)

    stat = database_path.stat()
    return {
        "schema_version": 1,
        "report_type": "company_profit_monthly",
        "month": month,
        "generated_at": datetime.now(
            ZoneInfo("Asia/Shanghai")
        ).isoformat(timespec="seconds"),
        "source": {
            "database_name": database_path.name,
            "database_size_bytes": stat.st_size,
            "database_modified_at": datetime.fromtimestamp(
                stat.st_mtime,
                ZoneInfo("Asia/Shanghai"),
            ).isoformat(timespec="seconds"),
            "mode": "sqlite_readonly_backup_then_offline_calculation",
        },
        "summary": summary,
    }


def write_private_json(snapshot, output_path):
    output_path = _validate_output_path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(
        snapshot,
        ensure_ascii=False,
        indent=2,
        sort_keys=False,
    ).encode("utf-8")
    descriptor = os.open(
        output_path,
        os.O_WRONLY | os.O_CREAT | os.O_TRUNC,
        0o600,
    )
    try:
        os.write(descriptor, payload)
        os.fsync(descriptor)
    finally:
        os.close(descriptor)
    try:
        os.chmod(output_path, 0o600)
    except OSError:
        pass
    return output_path


def main(argv=None):
    parser = ArgumentParser(
        description="从只读数据库快照生成离线公司经营月报数据"
    )
    parser.add_argument("--month", required=True, help="月份，格式 YYYY-MM")
    parser.add_argument(
        "--database",
        default=str(DEFAULT_DATABASE),
        help="SQLite 数据库路径",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="网站目录之外的私有 JSON 快照路径",
    )
    args = parser.parse_args(argv)

    snapshot = build_snapshot(args.month, args.database)
    output = write_private_json(snapshot, args.output)
    print(
        json.dumps(
            {
                "success": True,
                "month": snapshot["month"],
                "output": str(output),
                "actual_complete": snapshot["summary"]["actual"]["complete"],
                "forecast_complete": snapshot["summary"]["forecast"]["complete"],
                "data_gap_count": len(snapshot["summary"]["data_gaps"]),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
