import copy
import pytest

from shipment_retry import ShipmentRetryError, classify_retry_preflight


ITEMS = [{'id': 1179, 'product_id': 31263, 'variation_id': 31268, 'quantity': 1}]


def order(status='processing', tracking=None):
    return {'id': 33400, 'status': status, 'line_items': copy.deepcopy(ITEMS),
            'meta_data': [] if tracking is None else [{
                'key': '_wc_shipment_tracking_items',
                'value': [{'tracking_number': tracking}],
            }]}


def test_empty_tracking_on_unchanged_order_allows_explicit_retry():
    assert classify_retry_preflight(order(), '33400', ITEMS, 'ORIGINAL') == 'retry'


@pytest.mark.parametrize('status', ['processing', 'shipped', 'completed'])
def test_same_saved_parcel_only_reconciles(status):
    assert classify_retry_preflight(order(status, 'ORIGINAL'), '33400', ITEMS, 'ORIGINAL') == 'saved'


@pytest.mark.parametrize('change', ['id', 'quantity', 'product', 'missing_items', 'cancelled', 'refunded', 'other_tracking', 'completed_without_tracking'])
def test_changed_or_ambiguous_order_cannot_be_retried(change):
    remote = order()
    if change == 'id': remote['id'] = 33401
    elif change == 'quantity': remote['line_items'][0]['quantity'] = 2
    elif change == 'product': remote['line_items'][0]['variation_id'] = 31269
    elif change == 'missing_items': remote.pop('line_items')
    elif change in {'cancelled', 'refunded'}: remote['status'] = change
    elif change == 'other_tracking': remote = order(tracking='DIFFERENT')
    else: remote['status'] = 'completed'
    with pytest.raises(ShipmentRetryError):
        classify_retry_preflight(remote, '33400', ITEMS, 'ORIGINAL')


def test_multiple_saved_parcels_are_not_overwritten():
    remote = order(tracking='ORIGINAL')
    remote['meta_data'][0]['value'].append({'tracking_number': 'ANOTHER'})
    with pytest.raises(ShipmentRetryError):
        classify_retry_preflight(remote, '33400', ITEMS, 'ORIGINAL')


def test_item_order_does_not_change_the_preflight_result():
    items = ITEMS + [{'id': 1180, 'product_id': 29470, 'variation_id': 29714, 'quantity': 1}]
    remote = order()
    remote['line_items'] = list(reversed(items))
    assert classify_retry_preflight(remote, '33400', items, 'ORIGINAL') == 'retry'
