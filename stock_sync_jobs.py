"""Idempotent durable jobs and non-stealable physical-resource leases."""
from stock_sync_common import SyncError, one, rows, loads, dumps, digest, uid, stamp, now, parse_time, begin, lock_clause, event
from stock_sync_permissions import actor, require_object, target_sites
from stock_sync_catalog import enqueue

SUCCESS = {'verified_success','unchanged'}
RETRYABLE = {'failed','conflict','superseded','cancelled'}


def confirm(c,u,data):
    key = str(data.get('idempotency_key') or '')
    accepted = data.get('accepted_item_ids')
    if not key or len(key)>150 or not isinstance(accepted,list) or not accepted or any(not isinstance(x,str) for x in accepted):
        raise SyncError('INVALID_CONFIRMATION','确认需提交幂等键和明确项目列表',400)
    accepted = sorted(set(accepted))
    request_hash = digest({'plan_id':data.get('plan_id'),'version':data.get('plan_version'),'accepted':accepted})
    begin(c)
    # Serializes same-actor confirmations across different plans and keys.
    c.execute('SELECT id FROM users WHERE id=?'+lock_clause(c),(u['id'],)).fetchone()
    prior = one(c,'SELECT * FROM stock_sync_jobs WHERE actor_id=? AND idempotency_key=?',(u['id'],key))
    if prior:
        if prior['request_hash']!=request_hash:
            raise SyncError('IDEMPOTENCY_CONFLICT')
        c.commit();return prior['id']
    plan = one(c,'SELECT * FROM stock_sync_plans WHERE id=?'+lock_clause(c),(data.get('plan_id'),))
    if not plan:
        raise SyncError('NOT_FOUND',status=404)
    req = loads(plan['request_json'])
    require_object(c,u,plan,req['target_site_ids'],req.get('source_site_id'))
    prior = one(c,'SELECT * FROM stock_sync_jobs WHERE plan_id=?',(plan['id'],))
    if prior:
        if prior['request_hash'] != request_hash:
            raise SyncError('PLAN_CONSUMED')
        c.commit();return prior['id']
    if plan['status']!='ready' or data.get('plan_version')!=plan['version'] or parse_time(plan['expires_at'])<now():
        raise SyncError('PLAN_STALE')
    items = rows(c,'SELECT * FROM stock_sync_plan_items WHERE plan_id=?',(plan['id'],))
    chosen = [i for i in items if i['id'] in accepted]
    if len(chosen)!=len(accepted) or any(i['decision'] not in ('change','unchanged') for i in chosen):
        raise SyncError('INVALID_ACCEPTED_ITEMS','冲突和跳过项目不能执行',400)
    id_ = uid()
    c.execute("INSERT INTO stock_sync_jobs(id,plan_id,actor_id,idempotency_key,request_hash,status,created_at) VALUES(?,?,?,?,?,'queued',?)",(id_,plan['id'],u['id'],key,request_hash,stamp()))
    for item in chosen:
        c.execute("INSERT INTO stock_sync_job_items(id,job_id,plan_item_id,resource_key,status,updated_at) VALUES(?,?,?,?,'pending',?)",(uid(),id_,item['id'],item['resource_key'],stamp()))
    c.execute("UPDATE stock_sync_plans SET status='consumed' WHERE id=?",(plan['id'],))
    enqueue(c,'publish',id_)
    event(c,'plan_confirmed',u['id'],id_,{'plan_id':plan['id'],'accepted_item_ids':accepted})
    c.commit()
    return id_


def claim(c,worker_id):
    begin(c)
    suffix = ' FOR UPDATE SKIP LOCKED' if hasattr(c,'_raw') else ''
    work = one(c,"SELECT * FROM stock_sync_work WHERE status='queued' ORDER BY created_at,id LIMIT 1"+suffix)
    if not work:
        c.commit();return None
    c.execute("UPDATE stock_sync_work SET status='running',worker_id=?,heartbeat_at=? WHERE id=? AND status='queued'",(worker_id,stamp(),work['id']))
    c.commit()
    return work


def acquire(c,keys,owner):
    begin(c)
    for key in sorted(keys):
        cur = c.execute('''INSERT INTO stock_sync_resource_leases(resource_key,owner,version,heartbeat_at)
            VALUES(?,?,1,?) ON CONFLICT(resource_key) DO NOTHING''',(key,owner,stamp()))
        if cur.rowcount!=1:
            c.rollback()
            raise SyncError('RESOURCE_BUSY','资源正在执行或等待核实，请稍后重新预览')
    c.commit()


def release(c,owner):
    c.execute('DELETE FROM stock_sync_resource_leases WHERE owner=? AND quarantined=0',(owner,))
    c.commit()


def finish(c,job_id):
    job = one(c,'SELECT * FROM stock_sync_jobs WHERE id=?',(job_id,))
    counts = {r['status']:r['n'] for r in rows(c,'SELECT status,count(*) AS n FROM stock_sync_job_items WHERE job_id=? GROUP BY status',(job_id,))}
    states = set(counts)
    if states.intersection({'running','uncertain'}):
        status = 'requires_review'
    elif 'pending' in states:
        status = 'running'
    elif states.issubset(SUCCESS):
        status = 'succeeded'
    elif states=={'cancelled'}:
        status = 'cancelled'
    elif states.intersection(SUCCESS):
        status = 'partial_failed'
    else:
        status = 'cancelled' if job['cancel_requested'] else 'failed'
    c.execute('UPDATE stock_sync_jobs SET status=?,completed_at=? WHERE id=?',(status,stamp() if status!='running' else None,job_id))
    c.commit()
    return status


def cancel(c,u,job):
    plan = one(c,'SELECT * FROM stock_sync_plans WHERE id=?',(job['plan_id'],))
    req = loads(plan['request_json'])
    require_object(c,u,job,req['target_site_ids'],req.get('source_site_id'))
    c.execute("UPDATE stock_sync_jobs SET cancel_requested=1,status='cancel_requested' WHERE id=?",(job['id'],))
    c.execute("UPDATE stock_sync_job_items SET status='cancelled',updated_at=? WHERE job_id=? AND status='pending'",(stamp(),job['id']))
    event(c,'cancel_requested',u['id'],job['id'],{})
    c.commit()
    return finish(c,job['id'])
