"""Exercise the deployed clone functions with synthetic WooCommerce responses."""
import ast
import copy
import json
from pathlib import Path

import pytest
import requests

from product_clone_sku import build_clone_sku, make_clone_suffix, normalize_clone_suffix
from product_manager_service import parse_wc_response


class Response:
    def __init__(self, data, status=200):
        self.data = data
        self.status_code = status
        self.text = json.dumps(data)

    def json(self):
        return copy.deepcopy(self.data)


class WooFixture:
    def __init__(self):
        self.source = {'id': 100, 'name': 'Fixture', 'type': 'simple', 'sku': 'PARENT'}
        self.variations = [
            {'id': 101, 'sku': 'VARIANT', 'attributes': []},
            {'id': 102, 'sku': '', 'attributes': []},
        ]
        self.products = {'PARENT': {'id': 900, 'sku': 'PARENT', 'status': 'publish'}}
        self.created = []
        self.fail_variation_sku = False

    def get(self, url, **kwargs):
        if url == 'https://source.invalid/wp-json/wc/v3/products/100':
            return Response(self.source)
        if url == 'https://source.invalid/wp-json/wc/v3/products/100/variations':
            return Response(self.variations)
        assert url == 'https://target.invalid/wp-json/wc/v3/products'
        product = self.products.get(kwargs['params']['sku'])
        return Response([product] if product else [])

    def post(self, url, **kwargs):
        assert url.startswith('https://target.invalid/wp-json/wc/v3/products')
        payload = copy.deepcopy(kwargs['json'])
        self.created.append((url, payload))
        if url.endswith('/variations') and self.fail_variation_sku:
            return Response({'code': 'product_invalid_sku', 'message': 'SKU already exists'}, 400)
        result = {'id': 1000 + len(self.created), **payload}
        if not url.endswith('/variations'):
            self.products[payload['sku']] = result
        return Response(result)


@pytest.fixture
def clone(monkeypatch):
    fake = WooFixture()
    monkeypatch.setattr(requests.sessions.Session, 'request',
                        lambda *a, **kw: pytest.fail('Real HTTP forbidden'))
    monkeypatch.setattr(requests, 'get', fake.get)
    monkeypatch.setattr(requests, 'post', fake.post)
    scope = {
        '_WC_HEADERS': {'User-Agent': 'isolated-clone-test'},
        '_parse_wc_response': parse_wc_response,
        '_resolve_taxonomy_on_target': lambda *args: ([], []),
        'build_clone_sku': build_clone_sku,
        'make_clone_suffix': make_clone_suffix,
        'normalize_clone_suffix': normalize_clone_suffix,
    }
    # Importing app starts database initialization. Compile its real functions
    # here so these unit cases cannot initialize a configured production DB.
    tree = ast.parse((Path(__file__).resolve().parents[1] / 'app.py').read_text(encoding='utf-8'))
    functions = [n for n in tree.body if isinstance(n, ast.FunctionDef)
                 and n.name in {'_clone_one_product', '_clone_variations'}]
    assert len(functions) == 2
    exec(compile(ast.Module(body=functions, type_ignores=[]), 'app.py', 'exec'), scope)

    def run(mode='clone_as_new', **options):
        return scope['_clone_one_product'](
            'https://source.invalid', 'fixture', 'fixture',
            'https://target.invalid', 'fixture', 'fixture', 100,
            {'collision_mode': mode, 'clone_sku_suffix': 'NEW-TEST',
             'include_images': False, 'include_variations': True,
             'status_on_target': 'publish', **options})

    return fake, run


def test_default_clone_skips_existing_original_sku(clone):
    fake, run = clone
    result = run('skip_existing')
    assert result['new_id'] == 900
    assert result['skipped_existing'] is True
    assert not fake.created


def test_full_new_clone_creates_draft_despite_original_sku_collision(clone):
    fake, run = clone
    result = run()
    assert not result.get('skipped_existing')
    assert result['sku'] == 'PARENT-NEW-TEST'
    assert len(fake.created) == 1
    assert fake.created[0][1]['status'] == 'draft'
    assert fake.products['PARENT'] == {'id': 900, 'sku': 'PARENT', 'status': 'publish'}


def test_same_job_retry_reuses_new_sku_without_duplicate_creation(clone):
    fake, run = clone
    first = run()
    second = run()
    assert second['new_id'] == first['new_id']
    assert second['resumed_existing_clone'] is True
    assert second['skipped_existing'] is False
    assert len(fake.created) == 1


def test_existing_job_sku_is_reused_even_when_original_also_exists(clone):
    fake, run = clone
    fake.products['PARENT-NEW-TEST'] = {'id': 901, 'sku': 'PARENT-NEW-TEST', 'status': 'draft'}
    result = run()
    assert result['new_id'] == 901
    assert result['resumed_existing_clone'] is True
    assert not fake.created


def test_full_new_variable_clone_namespaces_parent_and_every_variation(clone):
    fake, run = clone
    fake.source['type'] = 'variable'
    result = run()
    assert not result.get('skipped_existing')
    assert [payload['sku'] for _, payload in fake.created] == [
        'PARENT-NEW-TEST', 'VARIANT-NEW-TEST', 'VAR-100-102-NEW-TEST']


def test_full_new_variant_collision_never_retries_without_sku(clone):
    fake, run = clone
    fake.source['type'] = 'variable'
    fake.variations = fake.variations[:1]
    fake.fail_variation_sku = True
    result = run()
    assert len(fake.created) == 2
    assert fake.created[1][1]['sku'] == 'VARIANT-NEW-TEST'
    assert any('1 失败' in warning for warning in result['warnings'])
