"""Build a bounded, repeatable restoration of an original legacy parcel.

Only order PUT integrations are eligible. AST's shipment POST is deliberately
excluded: replaying it can create another parcel or send another notification.
"""
import json
from datetime import datetime, timezone


MAX_REPAIR_ATTEMPTS = 3


class RepairNotAllowed(ValueError):
    pass


def restore_payload(snapshot, payload, remote, evidence, now):
    """Called only after full local validation and a fresh remote_absent GET."""
    previous = evidence.get('shipment_reconciliation', {})
    repair = evidence.get('shipment_repair', {})
    # Two independent reads, separated by the original grace period, establish
    # absence. A timeout/worker death never authorizes an immediate second PUT.
    if (previous.get('outcome') != 'remote_absent' or not previous.get('checked_at')
            or (now - datetime.fromisoformat(previous['checked_at'])).total_seconds() < 300):
        return None
    if int(repair.get('attempts', 0)) >= MAX_REPAIR_ATTEMPTS:
        raise RepairNotAllowed('原运单补同步已尝试 3 次仍未确认，请管理员检查站点接口')
    if evidence.get('format') not in {'villatheme', 'custom_lineitem'}:
        raise RepairNotAllowed('此站点的发货接口不支持自动补写，请管理员核对原运单')
    if (len(snapshot['logs']) != 1 or snapshot['logs'][0]['status'] != 'pending_sync'
            or snapshot['order']['delivery_confirmed']
            or snapshot['order']['status'] not in {'processing', 'on-hold'}
            or remote.get('status') not in {'processing', 'on-hold'}
            or payload.get('expected_status') != 'on-hold'):
        raise RepairNotAllowed('原发货记录或订单状态不满足补同步条件，请核对')
    # Do not replace damaged/unrecognized tracking blobs that the permissive
    # display extractor might have ignored. Empty metadata can be filled in.
    tracking_keys = {'_tracking_number', 'tracking_number', '_wc_shipment_tracking_items',
                     '_vi_wot_order_item_tracking_data'}
    for owner in [remote, *(remote.get('line_items') or []), *(remote.get('shipping_lines') or [])]:
        for meta in owner.get('meta_data') or []:
            if meta.get('key') in tracking_keys and meta.get('value') not in (None, '', [], '[]'):
                raise RepairNotAllowed('来源站点存在未识别的运单信息，请人工核对')
    timestamp = datetime.fromisoformat(str(snapshot['logs'][0]['shipped_at']))
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    shipped = int(timestamp.timestamp())
    slug, number = payload['carrier_slug'], payload['tracking_number']
    config = snapshot['carrier'] or {}
    url = (config.get('tracking_url') or '').replace('{tracking}', '{tracking_number}')

    def metadata(owner, key, value):
        matches = [m for m in owner.get('meta_data') or [] if m.get('key') == key]
        if len(matches) > 1:
            raise RepairNotAllowed('来源站点存在重复物流字段，请人工核对')
        result = {'key': key, 'value': value}
        if matches and matches[0].get('id'):
            result['id'] = matches[0]['id']
        return result

    result = {'status': 'on-hold', 'meta_data': [
        metadata(remote, '_tracking_number', number),
        metadata(remote, '_tracking_provider', slug),
        metadata(remote, '_date_shipped', str(shipped)),
    ], 'line_items': []}
    for line in remote['line_items']:
        if evidence['format'] == 'villatheme':
            value = json.dumps([{'tracking_number': number, 'carrier_slug': slug,
                                'carrier_name': config.get('name') or slug,
                                'carrier_url': url, 'carrier_type': 'custom-carrier',
                                'time': shipped}], ensure_ascii=False)
            values = [metadata(line, '_vi_wot_order_item_tracking_data', value)]
        else:
            values = [metadata(line, 'tracking_number', number),
                      metadata(line, 'carrier_slug', slug),
                      metadata(line, 'tracking_url', url.replace('{tracking_number}', number))]
        result['line_items'].append({'id': line['id'], 'meta_data': values})
    return result
