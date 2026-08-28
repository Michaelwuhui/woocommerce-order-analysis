import ast
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def _load_extract_custom_billing_fields():
    """Load the pure parser without importing the production Flask app."""
    source = ROOT.joinpath("app.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    function = next(
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef)
        and node.name == "extract_custom_billing_fields"
    )
    module = ast.Module(body=[function], type_ignores=[])
    namespace = {}
    exec(compile(module, str(ROOT / "app.py"), "exec"), namespace)
    return namespace["extract_custom_billing_fields"]


class HivapeInpostMetaTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.extract = staticmethod(_load_extract_custom_billing_fields())

    def test_hivape_full_locker_selection_is_displayed(self):
        fields = self.extract([
            {"key": "_paczkomat_id", "value": "DAB01BAPP"},
            {"key": "paczkomat_key", "value": "DAB01BAPP"},
            {
                "key": "Wybrany paczkomat",
                "value": "DAB01BAPP, Kolejowa 16, 62-069 Dąbrówka",
            },
        ])

        self.assertEqual(
            "DAB01BAPP, Kolejowa 16, 62-069 Dąbrówka",
            fields["customer_inpost_id"],
        )

    def test_hivape_locker_id_is_used_when_full_selection_is_missing(self):
        fields = self.extract([
            {"key": "paczkomat_key", "value": "OLD-FALLBACK"},
            {"key": "_paczkomat_id", "value": "ZGO41M"},
        ])

        self.assertEqual("ZGO41M", fields["customer_inpost_id"])

    def test_legacy_inpost_field_keeps_priority(self):
        fields = self.extract([
            {"key": "Wybrany paczkomat", "value": "HIVAPE-LOCKER, Address"},
            {"key": "_billing_inpost", "value": "LEGACY-LOCKER"},
        ])

        self.assertEqual("LEGACY-LOCKER", fields["customer_inpost_id"])

    def test_order_detail_uses_server_normalized_value(self):
        template = ROOT.joinpath("templates", "base.html").read_text(encoding="utf-8")
        self.assertIn(
            "var customerInPost = order.customer_inpost_id || '';",
            template,
        )


if __name__ == "__main__":
    unittest.main()
