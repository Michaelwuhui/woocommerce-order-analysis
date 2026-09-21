from copy import deepcopy
from datetime import datetime, timedelta, timezone
import json

import pytest

from shipment_repair import RepairNotAllowed, restore_payload
from test_shipment_reconciliation import parcel, source


def restoration():
    now = datetime(2026, 9, 21, tzinfo=timezone.utc)
    remote = source('none'); remote['status'] = 'processing'
    snapshot = {'order': {'status': 'processing', 'delivery_confirmed': False},
                'logs': [{'status': 'pending_sync', 'shipped_at': '2026-09-17 03:00:00'}],
                'carrier': {'name': 'InPost', 'tracking_url': 'https://tracking.invalid/{tracking}'}}
    payload = parcel(); payload['expected_status'] = 'on-hold'
    evidence = {'format': 'villatheme', 'shipment_reconciliation': {
        'outcome': 'remote_absent', 'checked_at': (now - timedelta(minutes=10)).isoformat()}}
    return snapshot, payload, remote, evidence, now


@pytest.mark.parametrize('fmt', ['villatheme', 'custom_lineitem'])
def test_patch_reuses_original_tracking_date_and_only_targets_metadata(fmt):
    snap, payload, remote, evidence, now = restoration()
    evidence['format'] = fmt
    remote['meta_data'] = [{'id': 901, 'key': '_tracking_number', 'value': ''},
                           {'id': 902, 'key': 'unrelated', 'value': 'keep'}]
    before = deepcopy(remote)
    patch = restore_payload(snap, payload, remote, evidence, now)
    assert patch == restore_payload(snap, payload, remote, evidence, now)
    assert patch['status'] == 'on-hold'
    assert patch['meta_data'][0] == {'id': 901, 'key': '_tracking_number', 'value': 'TRACK-123'}
    assert int(patch['meta_data'][2]['value']) == 1789614000
    assert remote == before and len(patch['line_items']) == 2
    for line in patch['line_items']:
        assert set(line) == {'id', 'meta_data'}
        if fmt == 'villatheme':
            record = json.loads(line['meta_data'][0]['value'])[0]
            assert record['time'] == 1789614000 and record['tracking_number'] == 'TRACK-123'
            assert record['carrier_url'] == 'https://tracking.invalid/{tracking_number}'


@pytest.mark.parametrize('change', ['ast', 'unknown', 'limit', 'delivered', 'completed', 'remote_cancelled',
                                   'missing_log', 'confirmed_log', 'wrong_target', 'damaged_tracking', 'duplicate_meta'])
def test_ambiguous_or_unsupported_restoration_is_not_automated(change):
    snap, payload, remote, evidence, now = restoration()
    if change in {'ast', 'unknown'}: evidence['format'] = change
    if change == 'limit': evidence['shipment_repair'] = {'attempts': 3}
    if change == 'delivered': snap['order']['delivery_confirmed'] = True
    if change == 'completed': snap['order']['status'] = 'completed'
    if change == 'remote_cancelled': remote['status'] = 'cancelled'
    if change == 'missing_log': snap['logs'] = []
    if change == 'confirmed_log': snap['logs'][0]['status'] = 'shipped'
    if change == 'wrong_target': payload['expected_status'] = 'shipped'
    if change == 'damaged_tracking':
        remote['line_items'][0]['meta_data'] = [{'key': '_vi_wot_order_item_tracking_data', 'value': 'broken'}]
    if change == 'duplicate_meta':
        remote['meta_data'] = [{'key': '_tracking_provider', 'value': ''}] * 2
    with pytest.raises(RepairNotAllowed): restore_payload(snap, payload, remote, evidence, now)


@pytest.mark.parametrize('previous', [{}, {'outcome': 'read_failed'}, {'outcome': 'needs_review'},
                                     {'outcome': 'remote_absent', 'checked_at': '2026-09-20T23:59:00+00:00'}])
def test_one_absence_or_recent_check_never_authorizes_write(previous):
    snap, payload, remote, evidence, now = restoration()
    evidence['shipment_reconciliation'] = previous
    assert restore_payload(snap, payload, remote, evidence, now) is None
