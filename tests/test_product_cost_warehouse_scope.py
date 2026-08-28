import ast
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class _Rows:
    def __init__(self, rows):
        self.rows = rows

    def fetchone(self):
        return self.rows[0] if self.rows else None

    def fetchall(self):
        return self.rows


class _Connection:
    def __init__(self, *, can_edit_costs, partner_ids=None):
        self.can_edit_costs = can_edit_costs
        self.partner_ids = partner_ids or []
        self.closed = False

    def execute(self, sql, params=()):
        if 'SELECT role, can_edit_costs FROM users' in sql:
            return _Rows([{'role': 'viewer', 'can_edit_costs': self.can_edit_costs}])
        if 'SELECT partner_id FROM partner_users' in sql:
            return _Rows([{'partner_id': value} for value in self.partner_ids])
        raise AssertionError(f'unexpected query: {sql}')

    def close(self):
        self.closed = True


class _User:
    is_authenticated = True
    username = 'internal-cost-editor'
    id = 7


def _load_scope_function(connection):
    source = ROOT.joinpath('app.py').read_text(encoding='utf-8')
    tree = ast.parse(source)
    function = next(
        node for node in tree.body
        if isinstance(node, ast.FunctionDef)
        and node.name == '_user_allowed_warehouse_ids'
    )
    namespace = {
        'current_user': _User(),
        'get_db_connection': lambda: connection,
    }
    exec(
        compile(ast.Module(body=[function], type_ignores=[]), str(ROOT / 'app.py'), 'exec'),
        namespace,
    )
    return namespace['_user_allowed_warehouse_ids']


class ProductCostWarehouseScopeTests(unittest.TestCase):
    def test_unbound_internal_cost_editor_can_manage_all_warehouses(self):
        connection = _Connection(can_edit_costs=1)

        self.assertIsNone(_load_scope_function(connection)())
        self.assertTrue(connection.closed)

    def test_unbound_user_without_cost_edit_grant_has_no_edit_scope(self):
        connection = _Connection(can_edit_costs=0)

        self.assertEqual([], _load_scope_function(connection)())
        self.assertTrue(connection.closed)

    def test_non_admin_does_not_see_admin_only_add_warehouse_button(self):
        template = ROOT.joinpath('templates', 'product_costs.html').read_text(encoding='utf-8')
        guard = '{% if current_user.is_admin() %}'
        button = 'onclick="quickAddWarehouse()"'

        self.assertIn(guard, template)
        self.assertLess(template.index(guard), template.index(button))
        self.assertLess(
            template.index(button),
            template.index('{% endif %}', template.index(guard)),
        )


if __name__ == '__main__':
    unittest.main()
