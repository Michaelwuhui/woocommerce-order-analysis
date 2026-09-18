import sqlite3
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import offline_company_profit_snapshot as snapshot_module


class OfflineCompanyProfitSnapshotTests(unittest.TestCase):
    def test_calculation_runs_inside_flask_application_context(self):
        class TrackingContext:
            active = False

            def __enter__(self):
                self.active = True
                return self

            def __exit__(self, exc_type, exc_value, traceback):
                self.active = False

        context = TrackingContext()
        app_module = SimpleNamespace(
            app=SimpleNamespace(app_context=lambda: context),
            get_db_connection=lambda: None,
            _compute_sales_board_data=lambda month: {},
            _revenue_status_cond=lambda: ("", []),
            _calc_partner_recon_detail=lambda *args, **kwargs: {},
            _compute_statement_split=lambda *args, **kwargs: {},
        )

        def require_context(*args, **kwargs):
            self.assertTrue(context.active)
            return {"actual": {}, "forecast": {}, "data_gaps": []}

        original_import_module = snapshot_module.importlib.import_module

        def import_app(name, package=None):
            if name == "app":
                return app_module
            return original_import_module(name, package)

        with tempfile.TemporaryDirectory() as tempdir:
            database_path = Path(tempdir) / "source.db"
            sqlite3.connect(database_path).close()
            with (
                patch.object(
                    snapshot_module.importlib,
                    "import_module",
                    side_effect=import_app,
                ),
                patch.object(
                    snapshot_module,
                    "init_company_profit_tables",
                    side_effect=require_context,
                ),
                patch.object(
                    snapshot_module,
                    "build_company_profit_summary",
                    side_effect=require_context,
                ),
            ):
                snapshot = snapshot_module.build_snapshot(
                    "2026-08",
                    database_path,
                )

        self.assertEqual(snapshot["month"], "2026-08")


if __name__ == "__main__":
    unittest.main()
