"""Run the actual cross-site editor controller with synthetic exact-site state."""
from pathlib import Path
import shutil
import subprocess

import pytest


ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.skipif(shutil.which("node") is None, reason="Node.js is required")
def test_cross_site_catalog_editor_contract():
    completed = subprocess.run(
        [shutil.which("node"), str(ROOT / "tests/product_manager_catalog_edit_harness.cjs")],
        cwd=ROOT, capture_output=True, text=True, encoding="utf-8", check=False,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr
