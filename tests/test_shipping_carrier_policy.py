import ast
import unittest
from pathlib import Path


class ShippingCarrierPolicyTests(unittest.TestCase):
    def test_polish_sites_require_inpost_and_dpd(self):
        source = Path(__file__).resolve().parents[1].joinpath("app.py").read_text(
            encoding="utf-8"
        )
        tree = ast.parse(source)
        policies = []
        for node in ast.walk(tree):
            if not isinstance(node, ast.Assign):
                continue
            if any(
                isinstance(target, ast.Name)
                and target.id == "required_slugs_by_country"
                for target in node.targets
            ):
                policies.append(ast.literal_eval(node.value))

        self.assertEqual(1, len(policies))
        self.assertEqual(("inpost", "dpd"), policies[0]["PL"])


if __name__ == "__main__":
    unittest.main()
