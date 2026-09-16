"""Operator maintenance: review a digest, then explicitly confirm the same state.

This CLI never writes WooCommerce. It requires DB access plus an active superadmin
actor. Independent-child evidence must come from a separately authorized site test.
"""
import argparse

from stock_sync_common import SyncError, connect, one, rows, loads, dumps, digest, stamp, uid, begin, event, lock_clause
from stock_sync_permissions import actor
from stock_sync_supply import mappings, mapping_hash
from stock_sync_woo import Woo, stock_state, site_identity, physical_scope, managed_resource_keys
from stock_sync_jobs import acquire, release


def superadmin(c,actor_id):
    u=actor(c,actor_id)
    if not u['superadmin']:
        raise SyncError('FORBIDDEN','该维护操作仅限超级管理员',403)
    return u


def capability(c,actor_id,site_id,evidence,enabled=True,confirmation=None):
    u=superadmin(c,actor_id)
    if not evidence or len(evidence.strip())<10:
        raise SyncError('EVIDENCE_REQUIRED','请记录子站独立写入及回读验证的证据位置')
    begin(c)
    site=one(c,'SELECT * FROM sites WHERE id=?'+lock_clause(c),(site_id,))
    if not site:raise SyncError('NOT_FOUND',status=404)
    current=one(c,'SELECT * FROM stock_sync_site_capabilities WHERE site_id=?',(site_id,))
    proof={'site_identity':site_identity(site),'evidence':evidence.strip()}
    state={'actor_id':u['id'],'site_id':site_id,'site':site['url'],'enabled':bool(enabled),'proof':proof,'current':current}
    token=digest(state)
    if confirmation is None:
        c.commit();return {'review':state,'confirm_hash':token}
    if confirmation!=token:raise SyncError('MAINTENANCE_STALE','配置或身份已改变，请重新核对')
    c.execute('''INSERT INTO stock_sync_site_capabilities(site_id,independent_write,evidence,version,actor_id,updated_at)
        VALUES(?,?,?,?,?,?) ON CONFLICT(site_id) DO UPDATE SET independent_write=excluded.independent_write,
        evidence=excluded.evidence,version=excluded.version,actor_id=excluded.actor_id,updated_at=excluded.updated_at''',
        (site_id,bool(enabled),dumps(proof),1 if not current else current['version']+1,u['id'],stamp()))
    event(c,'site_capability_changed',u['id'],str(site_id),state)
    c.commit();return {'saved':True,'site_id':site_id,'enabled':bool(enabled)}


def binding_preview(c,actor_id,map_id,reason,woo=None):
    u=superadmin(c,actor_id)
    if not reason or not reason.strip():raise SyncError('REASON_REQUIRED')
    mapping=next((m for m in mappings(c) if m['id']==map_id),None)
    old=one(c,'SELECT * FROM stock_sync_bindings WHERE map_id=?',(map_id,))
    if not mapping or not old:raise SyncError('NOT_FOUND','需要已接管且当前启用的映射')
    site=one(c,'SELECT * FROM sites WHERE id=?',(mapping['site_id'],))
    scope=physical_scope(c,site,mapping,[site['id']])
    old_baseline=loads(old['baseline_json'])
    if old['mapping_hash']==mapping_hash(mapping) and old_baseline.get('resource_key')==scope['resource_key']:
        raise SyncError('NO_BINDING_CHANGE')
    controls=rows(c,'SELECT * FROM stock_sync_controls WHERE map_id=? AND active=1 ORDER BY id',(map_id,))
    woo=woo or Woo();woo.connection=c;c.commit()
    remote=woo.read(site,mapping);woo.validate(remote,mapping)
    keys=sorted(set(old_baseline.get('managed_resource_keys',[])+managed_resource_keys(c,site,mapping)))
    preserves_quantity=old_baseline.get('mode')=='quantity' and bool(set(old_baseline.get('managed_resource_keys',[])).intersection(managed_resource_keys(c,site,mapping)))
    baseline={'mode':'quantity' if remote['manage_stock'] is True or preserves_quantity else 'status','state':stock_state(remote),'resource_key':scope['resource_key'],'managed_resource_keys':keys}
    review={'actor_id':u['id'],'mapping':mapping,'old_binding':old,'scope':scope,'baseline':baseline,
        'archive_controls':[x['id'] for x in controls],'controls':controls,'reason':reason.strip()}
    c.commit();return {'review':review,'confirm_hash':digest(review)}


def rebind(c,actor_id,map_id,reason,confirmation,woo=None):
    preview=binding_preview(c,actor_id,map_id,reason,woo)
    if preview['confirm_hash']!=confirmation:raise SyncError('MAINTENANCE_STALE')
    review=preview['review'];owner='maintenance:'+uid()
    keys=review['baseline']['managed_resource_keys']
    acquire(c,sorted(set(keys+['endpoint:'+k.split('/products/')[0] for k in keys])),owner)
    try:
        # Acquire both the historical and current physical resources, then repeat
        # the complete read. Rebinding is never a way to steal a worker's lease.
        current=binding_preview(c,actor_id,map_id,reason,woo)
        if current['confirm_hash']!=confirmation:raise SyncError('MAINTENANCE_STALE')
        begin(c)
        m=one(c,'SELECT * FROM inv_site_sku_map WHERE id=?'+lock_clause(c),(map_id,))
        if mapping_hash(m)!=mapping_hash(review['mapping']):raise SyncError('MAPPING_CHANGED')
        c.execute('''UPDATE stock_sync_controls SET active=0,version=version+1,released_by=?,released_at=?
            WHERE map_id=? AND active=1''',(actor_id,stamp(),map_id))
        c.execute('''UPDATE stock_sync_bindings SET site_id=?,sku_id=?,mapping_hash=?,baseline_json=?,
            version=version+1,updated_at=? WHERE map_id=?''',
            (m['site_id'],m['sku_id'],mapping_hash(m),dumps(review['baseline']),stamp(),map_id))
        event(c,'mapping_rebound',actor_id,str(map_id),review)
        c.commit()
        return {'saved':True,'map_id':map_id,'archived_controls':review['archive_controls'],
            'website_changed':False,'next_step':'重新生成差异预览；历史资源保护继续保留'}
    except Exception:
        c.rollback();raise
    finally:
        release(c,owner)


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--actor-id',type=int,required=True)
    sub=p.add_subparsers(dest='command',required=True)
    cap=sub.add_parser('capability')
    cap.add_argument('--site-id',type=int,required=True)
    cap.add_argument('--evidence',required=True)
    cap.add_argument('--disable',action='store_true')
    cap.add_argument('--confirm')
    binding=sub.add_parser('rebind')
    binding.add_argument('--map-id',type=int,required=True)
    binding.add_argument('--reason',required=True)
    binding.add_argument('--confirm')
    args=p.parse_args();c=connect()
    try:
        if args.command=='capability':
            result=capability(c,args.actor_id,args.site_id,args.evidence,not args.disable,args.confirm)
        elif args.confirm:
            result=rebind(c,args.actor_id,args.map_id,args.reason,args.confirm)
        else:
            result=binding_preview(c,args.actor_id,args.map_id,args.reason)
        print(dumps(result))
    except SyncError as exc:
        c.rollback();p.exit(2,exc.code+': '+str(exc)+'\n')
    finally:c.close()


if __name__=='__main__':
    main()
