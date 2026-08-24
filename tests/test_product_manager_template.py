import re
import shutil
import subprocess
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
PRODUCT_MANAGER_TEMPLATE = ROOT / "templates" / "product_manager.html"


def product_manager_script():
    template = PRODUCT_MANAGER_TEMPLATE.read_text(encoding="utf-8")
    scripts = re.findall(r"<script(?:\s[^>]*)?>([\s\S]*?)</script>", template)
    assert len(scripts) == 1
    return scripts[0]


def test_product_manager_clone_resume_block_is_not_duplicated():
    script = product_manager_script()
    assert script.count("const activeCloneJob =") == 1
    assert "$('pmLoadBtn').addEventListener('click', () => loadProducts(1));" in script


@pytest.mark.skipif(shutil.which("node") is None, reason="Node.js is required")
def test_product_manager_inline_script_has_valid_javascript_syntax():
    completed = subprocess.run(
        [shutil.which("node"), "--check", "-"],
        input=product_manager_script(),
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
