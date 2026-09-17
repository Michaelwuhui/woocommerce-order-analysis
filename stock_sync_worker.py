"""Standalone worker; closed browser tabs never interrupt committed work.

Leases never expire into a second write. After a crashed worker is known stopped,
use --recover-worker ID --confirm-worker-stopped for READ-ONLY reconciliation.
"""
import argparse
import os
import socket
import threading
import time

from stock_sync_common import SyncError, connect, one, rows, loads, dumps, digest, uid, stamp, now, parse_time, begin, event, enabled
from stock_sync_permissions import actor, require_object
from stock_sync_supply import mapping_hash, mappings
from stock_sync_policy import matches
from stock_sync_catalog import scan
from stock_sync_planner import build_plan, evaluate
from stock_sync_woo import Woo, stock_state
from stock_sync_jobs import claim, acquire, release, finish, SUCCESS
import stock_sync_mapping as mapping_assist


def control_hash(c,map_id):
    return digest(rows(c,'SELECT * FROM stock_sync_controls WHERE map_id=? AND active=1 ORDER BY created_at,id',(map_id,)))


def check_intent(c,job,request,detail,intent):
    u = actor(c,job['actor_id'])
    require_object(c,u,job,request['target_site_ids'],request.get('source_site_id'))
    m = one(c,'SELECT * FROM inv_site_sku_map WHERE id=? AND is_active=1',(detail['map_id'],))
    if not m or mapping_hash(m)!=intent['mapping_hash']:
        raise SyncError('MAPPING_CHANGED')
    if control_hash(c,detail['map_id']) != intent['controls_hash']:
        raise SyncError('SUPERSEDED')
    lease = one(c,'SELECT * FROM stock_sync_resource_leases WHERE resource_key=?',(detail['resource_key'],))
    if not lease or lease['owner']!=intent['owner'] or lease['quarantined']:
        raise SyncError('RESOURCE_BUSY')
    c.commit()


def persist_intent(c,job,item,detail,request):
    begin(c)
    # A lease serializes control changes as well as remote writes. Recheck the
    # current DB actor after acquiring it, never trust the queued user snapshot.
    u = actor(c,job['actor_id'])
    require_object(c,u,job,request['target_site_ids'],request.get('source_site_id'))
    if one(c,'SELECT cancel_requested FROM stock_sync_jobs WHERE id=?',(job['id'],))['cancel_requested']:
        raise SyncError('CANCELLED')
    mapping = one(c,'SELECT * FROM inv_site_sku_map WHERE id=?',(detail['map_id'],))
    if mapping_hash(mapping)!=mapping_hash(detail['mapping']) or control_hash(c,detail['map_id'])!=digest(detail['controls']):
        raise SyncError('SUPERSEDED')
    old = one(c,'SELECT * FROM stock_sync_bindings WHERE map_id=?',(detail['map_id'],))
    c.execute('''INSERT INTO stock_sync_bindings(map_id,site_id,sku_id,mapping_hash,baseline_json,policy,version,updated_at)
        VALUES(?,?,?,?,?,'stock_sync',?,?) ON CONFLICT(map_id) DO UPDATE SET
        version=excluded.version,updated_at=excluded.updated_at''',
        (detail['map_id'],detail['site_id'],detail['sku_id'],mapping_hash(mapping),dumps(detail['baseline']),1 if not old else old['version']+1,stamp()))
    for index,change in enumerate(detail['changes']):
        if 'release' in change:
            c.execute('UPDATE stock_sync_controls SET active=0,version=version+1,released_by=?,released_at=? WHERE id=? AND active=1', (u['id'],stamp(),change['release']))
        else:
            if change['kind'] in ('reference_snapshot','manual_availability'):
                c.execute('UPDATE stock_sync_controls SET active=0,version=version+1,released_by=?,released_at=? WHERE map_id=? AND kind=? AND active=1',(u['id'],stamp(),detail['map_id'],change['kind']))
            c.execute('''INSERT INTO stock_sync_controls(id,map_id,site_id,sku_id,mapping_hash,kind,stock_status,
                protection,source_json,reason,active,version,actor_id,created_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,1,1,?,?)''',
                (f'{item["id"]}:{index}',detail['map_id'],detail['site_id'],detail['sku_id'],mapping_hash(mapping),
                 change['kind'],change['stock_status'],change['protection'],dumps(change.get('source',{})),request['reason'],u['id'],stamp()))
    intent = {'owner':item['id'],'mapping_hash':mapping_hash(mapping),'controls_hash':control_hash(c,detail['map_id']),
              'intended':detail['intended'],'controls_saved':True}
    c.execute("UPDATE stock_sync_job_items SET status='running',intent_json=?,before_json=?,updated_at=? WHERE id=?",(dumps(intent),dumps(detail['before']),stamp(),item['id']))
    event(c,'publish_intent',u['id'],item['id'],{'changes':detail['changes'],'intended':detail['intended']})
    c.commit()
    return intent


def save_result(c,job,item,status,error,after=None):
    c.execute('UPDATE stock_sync_job_items SET status=?,error=?,after_json=?,updated_at=? WHERE id=?',
              (status,error,dumps(stock_state(after)) if after else None,stamp(),item['id']))
    event(c,'readback' if after else 'item_stopped',job['actor_id'],item['id'],{'status':status,'error':error,'after':stock_state(after) if after else None})
    c.commit()


def process_item(c,job,plan,item,woo):
    detail = loads(one(c,'SELECT detail_json FROM stock_sync_plan_items WHERE id=?',(item['plan_item_id'],))['detail_json'])
    request = loads(plan['request_json'])
    lease = False
    intent = None
    try:
        # Endpoint-level lease additionally limits this adapter to one request per
        # site. Different logical aliases of an endpoint use the same key.
        endpoint = detail['resource_key'].split('/products/')[0]
        acquire(c,[detail['resource_key'],'endpoint:'+endpoint],item['id'])
        lease = True
        if parse_time(plan['expires_at']) < now():
            raise SyncError('PLAN_STALE')
        u = actor(c,job['actor_id'])
        current = next((m for m in mappings(c,[detail['site_id']]) if m['id']==detail['map_id']),None)
        if not current:
            raise SyncError('MAPPING_CHANGED')
        checked = evaluate(c,u,request,current,woo)
        if checked['inputs_hash']!=detail['inputs_hash']:
            raise SyncError('REMOTE_STATE_CHANGED')
        site = one(c,'SELECT * FROM sites WHERE id=?',(detail['site_id'],))
        c.commit()
        before = woo.read(site,current)
        if stock_state(before)!=detail['before']:
            raise SyncError('REMOTE_STATE_CHANGED')
        intent = persist_intent(c,job,item,detail,request)
        if matches(before,intent['intended']):
            check_intent(c,job,request,detail,intent)
            save_result(c,job,item,'unchanged',None,before)
            return
        c.execute('UPDATE stock_sync_job_items SET attempts=attempts+1,updated_at=? WHERE id=?',(stamp(),item['id']))
        event(c,'write_attempt',job['actor_id'],item['id'],{'payload':intent['intended']})
        c.commit()
        write_error = None
        try:
            woo.publish(site,current,before,intent['intended'],lambda:check_intent(c,job,request,detail,intent))
        except SyncError as exc:
            write_error = exc.code
        if write_error=='REMOTE_RATE_LIMITED':
            save_result(c,job,item,'failed','REMOTE_RATE_LIMITED')
            return
        # Even a timed-out PUT is followed by GET; it is never blindly replayed.
        try:
            after = woo.read(site,current)
        except SyncError:
            save_result(c,job,item,'uncertain','WRITE_RESULT_UNKNOWN')
            c.execute('UPDATE stock_sync_resource_leases SET quarantined=1 WHERE owner=?',(item['id'],));c.commit()
            return
        check_intent(c,job,request,detail,intent)
        if matches(after,intent['intended']):
            save_result(c,job,item,'verified_success','VERIFIED_AFTER_ERROR' if write_error else None,after)
        elif write_error=='WRITE_RESULT_UNKNOWN':
            save_result(c,job,item,'uncertain','WRITE_RESULT_UNKNOWN',after)
            c.execute('UPDATE stock_sync_resource_leases SET quarantined=1 WHERE owner=?',(item['id'],));c.commit()
        else:
            save_result(c,job,item,'failed',write_error or 'READBACK_MISMATCH',after)
    except SyncError as exc:
        c.rollback()
        status = 'cancelled' if exc.code=='CANCELLED' else ('superseded' if exc.code=='SUPERSEDED' else 'conflict')
        # An error after intent may follow a write; retain the lease and require
        # reconciliation rather than permit a newer job to race that request.
        if intent:
            status = 'uncertain'
            c.execute('UPDATE stock_sync_resource_leases SET quarantined=1 WHERE owner=?',(item['id'],));c.commit()
        save_result(c,job,item,status,exc.code)
    except Exception:
        c.rollback()
        # If result persistence failed, the intent stays committed. Do not remove
        # its lease; recovery needs it to establish what actually happened.
        if lease:
            c.execute('UPDATE stock_sync_resource_leases SET quarantined=1 WHERE owner=?',(item['id'],));c.commit()
        raise
    finally:
        if lease:
            release(c,item['id'])


def process_job(c,id_,woo=None):
    woo = woo or Woo()
    woo.connection = c
    job = one(c,'SELECT * FROM stock_sync_jobs WHERE id=?',(id_,))
    plan = one(c,'SELECT * FROM stock_sync_plans WHERE id=?',(job['plan_id'],))
    c.execute("UPDATE stock_sync_jobs SET status='running' WHERE id=? AND cancel_requested=0",(id_,));c.commit()
    for item in rows(c,"SELECT * FROM stock_sync_job_items WHERE job_id=? AND status='pending' ORDER BY id",(id_,)):
        if one(c,'SELECT cancel_requested FROM stock_sync_jobs WHERE id=?',(id_,))['cancel_requested']:
            c.execute("UPDATE stock_sync_job_items SET status='cancelled',updated_at=? WHERE id=? AND status='pending'",(stamp(),item['id']));c.commit()
            continue
        process_item(c,job,plan,item,woo)
    return finish(c,id_)


def recover(c,worker_id,woo=None):
    """Read-only reconciliation after the operator confirms the worker is gone.

    Recovery does not need the original user's current write permission: it never
    writes a website or creates controls. Future retry plans do recheck permission.
    """
    woo = woo or Woo()
    woo.connection = c
    work = rows(c,"SELECT * FROM stock_sync_work WHERE worker_id=? AND status='running'",(worker_id,))
    for w in work:
        if w['kind']!='publish':
            table = 'stock_sync_plans' if w['kind']=='plan' else 'stock_sync_catalog_snapshots'
            c.execute(f"UPDATE {table} SET status='failed',error='WORKER_INTERRUPTED' WHERE id=?",(w['object_id'],))
        else:
            job = one(c,'SELECT * FROM stock_sync_jobs WHERE id=?',(w['object_id'],))
            items = rows(c,"SELECT * FROM stock_sync_job_items WHERE job_id=? AND status NOT IN ('verified_success','unchanged')",(job['id'],))
            for item in items:
                if not item['intent_json']:
                    save_result(c,job,item,'conflict','WORKER_INTERRUPTED')
                else:
                    detail = loads(one(c,'SELECT detail_json FROM stock_sync_plan_items WHERE id=?',(item['plan_item_id'],))['detail_json'])
                    intent = loads(item['intent_json'])
                    site = one(c,'SELECT * FROM sites WHERE id=?',(detail['site_id'],))
                    mapping = one(c,'SELECT * FROM inv_site_sku_map WHERE id=?',(detail['map_id'],))
                    try:
                        if not mapping or mapping_hash(mapping)!=intent['mapping_hash'] or control_hash(c,detail['map_id'])!=intent['controls_hash']:
                            save_result(c,job,item,'superseded','SUPERSEDED')
                        else:
                            c.commit()
                            after = woo.read(site,mapping)
                            save_result(c,job,item,'verified_success' if matches(after,intent['intended']) else 'failed','RECOVERED_READBACK',after)
                    except SyncError:
                        save_result(c,job,item,'uncertain','WRITE_RESULT_UNKNOWN')
                        continue
                c.execute('UPDATE stock_sync_resource_leases SET quarantined=0 WHERE owner=?',(item['id'],));c.commit()
                release(c,item['id'])
            finish(c,job['id'])
        # Keep unresolved publish jobs discoverable by a subsequent recovery.
        unresolved = w['kind']=='publish' and one(c,"SELECT id FROM stock_sync_job_items WHERE job_id=? AND status='uncertain' LIMIT 1",(w['object_id'],))
        if not unresolved:
            c.execute("UPDATE stock_sync_work SET status='recovered' WHERE id=?",(w['id'],))
        c.commit()


def run_once(c,worker_id,woo=None):
    work = claim(c,worker_id)
    if not work:
        return False
    try:
        {'catalog':scan,'plan':build_plan,'publish':process_job,
         'mapping':mapping_assist.scan,'mapping_confirm':mapping_assist.apply}[work['kind']](c,work['object_id'],woo)
        unresolved = work['kind']=='publish' and one(c,"SELECT id FROM stock_sync_job_items WHERE job_id=? AND status IN ('running','uncertain') LIMIT 1",(work['object_id'],))
        if not unresolved:
            c.execute("UPDATE stock_sync_work SET status='done',heartbeat_at=? WHERE id=?",(stamp(),work['id']));c.commit()
    except Exception:
        c.rollback()
        if work['kind']=='publish':
            finish(c,work['object_id'])
        raise
    return True


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--once',action='store_true')
    p.add_argument('--recover-worker')
    p.add_argument('--confirm-worker-stopped',action='store_true')
    args=p.parse_args()
    if not enabled():
        p.error('STOCK_SYNC_ENABLED=1 is required')
    worker_id=f'{socket.gethostname()}:{os.getpid()}:{uid()[:8]}'
    print('stock_sync_worker='+worker_id,flush=True)
    c=connect()
    if args.recover_worker:
        if not args.confirm_worker_stopped:
            p.error('Stop the previous process, then pass --confirm-worker-stopped')
        recover(c,args.recover_worker);c.close();return
    stop=threading.Event()
    def heartbeat():
        while not stop.wait(10):
            h=connect()
            try:
                h.execute("UPDATE stock_sync_work SET heartbeat_at=? WHERE worker_id=? AND status='running'",(stamp(),worker_id))
                h.execute('''UPDATE stock_sync_resource_leases SET heartbeat_at=? WHERE owner IN
                    (SELECT i.id FROM stock_sync_job_items i JOIN stock_sync_work w ON w.object_id=i.job_id
                     WHERE w.worker_id=? AND w.status='running' AND i.status='running')''',(stamp(),worker_id))
                h.commit()
            finally:
                h.close()
    t=threading.Thread(target=heartbeat,daemon=True);t.start()
    try:
        while True:
            worked=run_once(c,worker_id)
            if args.once: break
            if not worked: time.sleep(1)
    finally:
        stop.set();c.close()


if __name__=='__main__':
    main()
