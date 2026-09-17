from copy import deepcopy
import json

import pytest

from shipment_reconciliation import ReviewRequired, products, validate_local, verify_remote
from external_operations import canonical_hash


def parcel():
    return {'order_id': '20-123', 'site_id': 20, 'tracking_number': 'TRACK-123',
            'carrier_slug': 'inpost', 'items': [
                {'item_id': '11', 'product': '101', 'qty': '2'},
                {'item_id': '12', 'product': '102', 'qty': '1'}]}


def source(fmt='ast'):
    p = parcel()
    order = {'id': 123, 'status': 'on-hold', 'date_modified': '2026-09-17T10:00:00',
             'line_items': [{'id': 11, 'product_id': 100, 'variation_id': 101, 'quantity': 2, 'meta_data': []},
                            {'id': 12, 'product_id': 102, 'quantity': 1, 'meta_data': []}],
             'meta_data': [], 'shipping_lines': []}
    if fmt == 'ast':
        order['meta_data'] = [{'key': '_wc_shipment_tracking_items', 'value': [{
            'tracking_number': p['tracking_number'], 'tracking_provider': 'inpost-paczkomaty',
            'products_list': deepcopy(p['items'])}]}]
    elif fmt == 'villa':
        for line in order['line_items']:
            line['meta_data'] = [{'key': '_vi_wot_order_item_tracking_data', 'value': json.dumps([{
                'tracking_number': p['tracking_number'], 'carrier_slug': 'inpost'}])}]
    elif fmt == 'custom':
        for line in order['line_items']:
            line['meta_data'] = [{'key': 'tracking_number', 'value': p['tracking_number']},
                                 {'key': 'carrier_slug', 'value': 'inpost'}]
    return order


@pytest.mark.parametrize('fmt', ['ast', 'villa', 'custom'])
def test_exact_original_parcel_all_supported_formats(fmt):
    assert verify_remote(source(fmt), parcel()) == 'verified'


@pytest.mark.parametrize('change', ['wrong_order', 'wrong_variant', 'wrong_qty', 'other_tracking',
                                   'wrong_carrier', 'partial_ast', 'duplicate_ast', 'status_drift'])
def test_conflicting_remote_evidence_never_confirms(change):
    r = source()
    tracking = r['meta_data'][0]['value'][0]
    if change == 'wrong_order': r['id'] = 124
    if change == 'wrong_variant': r['line_items'][0]['variation_id'] = 999
    if change == 'wrong_qty': r['line_items'][0]['quantity'] = 1
    if change == 'other_tracking': tracking['tracking_number'] = 'OTHER'
    if change == 'wrong_carrier': tracking['tracking_provider'] = 'dpd'
    if change == 'partial_ast': tracking['products_list'][0]['qty'] = '1'
    if change == 'duplicate_ast': r['meta_data'][0]['value'].append(deepcopy(tracking))
    if change == 'status_drift': r['status'] = 'processing'
    with pytest.raises(ReviewRequired): verify_remote(r, parcel())


@pytest.mark.parametrize('fmt', ['villa', 'custom'])
def test_tracking_on_only_one_line_is_not_full_shipment(fmt):
    r = source(fmt)
    r['line_items'][1]['meta_data'] = []
    with pytest.raises(ReviewRequired): verify_remote(r, parcel())


def test_ast_tracking_without_product_coverage_cannot_confirm():
    r = source()
    r['meta_data'][0]['value'][0].pop('products_list')
    with pytest.raises(ValueError): verify_remote(r, parcel())


def test_unrecorded_tracking_is_not_success_or_a_retry_authorization():
    r = source('none'); r['status'] = 'processing'
    assert verify_remote(r, parcel()) == 'remote_absent'
    r['status'] = 'completed'
    with pytest.raises(ReviewRequired): verify_remote(r, parcel())


def test_duplicate_item_quantities_cannot_hide_in_a_parcel():
    with pytest.raises(ValueError): products(parcel()['items'] * 2)


def test_local_request_hash_and_full_parcel_are_required():
    p = parcel()
    snapshot = {'op': {'order_id': p['order_id'], 'site_id': 20, 'request_payload': p,
                       'request_hash': canonical_hash(p)},
                'order': {'source': 'https://test.invalid', 'status': 'processing',
                          'line_items': source()['line_items'], 'is_undelivered': False,
                          'is_problem_return': False},
                'site': {'id': 20, 'url': 'https://test.invalid'}, 'logs': [], 'has_oms': False}
    assert validate_local(snapshot) == p
    for flag in ('new_parcel', 'more_batches', 'is_reship'):
        bad = deepcopy(snapshot); bad['op']['request_payload'][flag] = True
        bad['op']['request_hash'] = canonical_hash(bad['op']['request_payload'])
        with pytest.raises(ReviewRequired): validate_local(bad)
    snapshot['has_oms'] = True
    with pytest.raises(ReviewRequired): validate_local(snapshot)


def test_source_metadata_does_not_leak_into_review_error():
    r = source(); r['meta_data'][0]['value'][0]['products_list'] = 'not valid json containing secrets'
    with pytest.raises(ValueError): verify_remote(r, parcel())
