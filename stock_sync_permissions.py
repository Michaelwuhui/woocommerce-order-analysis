"""Current DB permissions at every boundary, independent of the broad admin role."""
from stock_sync_common import SyncError, one, rows, positive_ids


def actor(c, actor_id):
    u = one(c, 'SELECT * FROM users WHERE id=?', (actor_id,))
    if not u or u.get('is_active', 1) != 1 or u.get('is_enabled', 1) != 1:
        raise SyncError('PERMISSION_REVOKED', '账户不存在或已停用', 403)
    u['superadmin'] = u.get('username') == 'admin'
    if not u['superadmin'] and not u.get('can_manage_products'):
        raise SyncError('PERMISSION_REVOKED', '无产品管理权限', 403)
    return u


def can_write(u, site):
    return u['superadmin'] or bool(u.get('name') and str(u['name']).strip() == str(site.get('manager') or '').strip())


def target_sites(c, u, scope):
    sites = rows(c, 'SELECT * FROM sites ORDER BY id')
    if scope.get('mode') == 'explicit_sites':
        ids = positive_ids(scope.get('site_ids'), 'site_ids')
        chosen = [s for s in sites if s['id'] in ids]
        if len(chosen) != len(ids) or any(not can_write(u, s) for s in chosen):
            raise SyncError('PERMISSION_REVOKED', '包含无权写入或不存在的目标站点', 403)
    elif scope.get('mode') == 'all_authorized':
        chosen = [s for s in sites if can_write(u, s)]
    else:
        raise SyncError('INVALID_INPUT', '请选择明确目标站点', 400)
    for key, field in [('managers', 'manager'), ('countries', 'country')]:
        if scope.get(key):
            if not isinstance(scope[key], list):
                raise SyncError('INVALID_INPUT', '筛选条件必须是列表', 400)
            chosen = [s for s in chosen if s.get(field) in scope[key]]
    return chosen


def reference_site(c, u, site_id):
    s = one(c, 'SELECT * FROM sites WHERE id=?', (site_id,))
    shared = one(c, 'SELECT * FROM stock_sync_reference_sites WHERE site_id=? AND enabled=1', (site_id,))
    if not s or not (can_write(u, s) or shared):
        raise SyncError('REFERENCE_PERMISSION_REVOKED', '参照读取权限不存在或已撤销', 403)
    return s


def visible_site(s):
    return {k: s.get(k) for k in ('id', 'url', 'name', 'manager', 'country')}


def can_release(u, control):
    return u['superadmin'] or control['protection'] != 'superadmin'


def require_object(c, u, obj, site_ids, source_id=None):
    if not u['superadmin'] and obj['actor_id'] != u['id']:
        raise SyncError('FORBIDDEN', '无权查看该记录', 403)
    target_sites(c, u, {'mode': 'explicit_sites', 'site_ids': list(site_ids)})
    if source_id:
        reference_site(c, u, source_id)
