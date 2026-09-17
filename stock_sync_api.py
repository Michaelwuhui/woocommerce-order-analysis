"""Stock-sync UI/API. Session authentication + dedicated CSRF protection."""
from copy import deepcopy
from functools import wraps
import secrets

from flask import Blueprint, jsonify, request, session, render_template
from flask_login import current_user, login_required

from stock_sync_common import SyncError, connect, one, rows, loads, dumps, stamp, uid, exists, enabled, event, begin, lock_clause
from stock_sync_permissions import actor, target_sites, reference_site, visible_site, require_object, can_write
from stock_sync_catalog import create_scan, enqueue
from stock_sync_planner import create_plan
from stock_sync_jobs import confirm, cancel, RETRYABLE

bp=Blueprint('stock_sync',__name__)
PREFIX='/api/product-manager/stock-sync'


def csrf_token():
    if 'stock_sync_csrf' not in session:
        session['stock_sync_csrf']=secrets.token_urlsafe(32)
    return session['stock_sync_csrf']


def api(func):
    @wraps(func)
    @login_required
    def wrapped(*args,**kwargs):
        if not enabled():
            return jsonify(error='库存同步尚未启用',code='FEATURE_DISABLED'),503
        c=connect()
        try:
            if not exists(c,'stock_sync_schema_migrations'):
                raise SyncError('SCHEMA_NOT_READY','库存同步迁移尚未安装',503)
            u=actor(c,int(current_user.id))
            if request.method not in ('GET','HEAD'):
                token=request.headers.get('X-CSRF-Token','')
                if not token or not secrets.compare_digest(token,session.get('stock_sync_csrf','')):
                    raise SyncError('CSRF_REJECTED','页面已过期，请刷新后重试',403)
                if not request.is_json:
                    raise SyncError('JSON_REQUIRED',status=400)
                if not isinstance(request.get_json(silent=True),dict):
                    raise SyncError('INVALID_INPUT','请求必须是 JSON 对象',400)
            return func(c,u,*args,**kwargs)
        except SyncError as exc:
            c.rollback()
            return jsonify(error=str(exc),code=exc.code),exc.status
        except (TypeError,ValueError,KeyError):
            c.rollback()
            return jsonify(error='请求格式不正确',code='INVALID_INPUT'),400
        finally:
            c.close()
    return wrapped


def accessible_plan(c,u,id_):
    p=one(c,'SELECT * FROM stock_sync_plans WHERE id=?',(id_,))
    if not p: raise SyncError('NOT_FOUND',status=404)
    req=loads(p['request_json'])
    require_object(c,u,p,req['target_site_ids'],req.get('source_site_id'))
    return p,req


def accessible_job(c,u,id_):
    j=one(c,'SELECT * FROM stock_sync_jobs WHERE id=?',(id_,))
    if not j: raise SyncError('NOT_FOUND',status=404)
    p,req=accessible_plan(c,u,j['plan_id'])
    if j['actor_id']!=u['id'] and not u['superadmin']:
        raise SyncError('FORBIDDEN',status=403)
    return j,p,req


def public_detail(u,d):
    d=deepcopy(d)
    d.pop('inputs_hash',None)
    for source in d.get('source',[]):
        source.pop('site_identity',None)
    if 'scope' in d:
        d['scope'].pop('site_identity',None)
        d['scope'].pop('capability',None)
    if not u['superadmin']:
        d.pop('mapping',None)
        if 'baseline' in d:
            d['baseline'].pop('managed_resource_keys',None)
        for source in d.get('source',[]):
            source.pop('pools',None)
        for control in d.get('controls',[]):
            control.pop('source_json',None)
        if 'supply' in d:
            d['supply']={'result':d['supply']['result']}
    return d


@bp.route('/product-manager/stock-sync')
@login_required
def page():
    c=connect()
    try:
        actor(c,int(current_user.id))
    except SyncError:
        return '无产品管理权限',403
    finally:
        c.close()
    return render_template('stock_sync.html',stock_sync_token=csrf_token(),stock_sync_enabled=enabled())


@bp.route(PREFIX+'/options')
@api
def options(c,u):
    all_sites=rows(c,'SELECT * FROM sites ORDER BY id')
    targets=[s for s in all_sites if can_write(u,s)]
    refs=[]
    for s in all_sites:
        try:
            reference_site(c,u,s['id'])
            refs.append(visible_site(s))
        except SyncError:
            pass
    return jsonify(target_sites=[{**visible_site(s),'available':bool(s.get('consumer_key') and s.get('consumer_secret') and s.get('is_active',1))} for s in targets],
        reference_sites=refs,superadmin=u['superadmin'],csrf_token=csrf_token(),
        reference_settings=rows(c,'SELECT site_id,enabled,version FROM stock_sync_reference_sites') if u['superadmin'] else [],
        worker=rows(c,"SELECT worker_id,heartbeat_at,kind FROM stock_sync_work WHERE status='running' ORDER BY created_at DESC LIMIT 3"))


@bp.route(PREFIX+'/reference-sites/<int:site_id>',methods=['PUT'])
@api
def set_reference(c,u,site_id):
    if not u['superadmin']: raise SyncError('FORBIDDEN',status=403)
    target_sites(c,u,{'mode':'explicit_sites','site_ids':[site_id]})
    value=request.get_json().get('enabled')
    if type(value) is not bool: raise SyncError('INVALID_INPUT',status=400)
    c.execute('''INSERT INTO stock_sync_reference_sites(site_id,enabled,version,actor_id,updated_at) VALUES(?,?,1,?,?) ON CONFLICT(site_id)
        DO UPDATE SET enabled=excluded.enabled,version=stock_sync_reference_sites.version+1,
        actor_id=excluded.actor_id,updated_at=excluded.updated_at''',(site_id,int(value),u['id'],stamp()))
    event(c,'reference_sharing',u['id'],str(site_id),{'enabled':value});c.commit()
    return jsonify(success=True)


@bp.route(PREFIX+'/catalog-scans',methods=['POST'])
@api
def scans(c,u):
    return jsonify(id=create_scan(c,u,request.get_json())),202


def accessible_scan(c,u,id_):
    snap=one(c,'SELECT * FROM stock_sync_catalog_snapshots WHERE id=?',(id_,))
    if not snap: raise SyncError('NOT_FOUND',status=404)
    scope=loads(snap['scope_json'])
    require_object(c,u,snap,scope['site_ids'],snap['site_id'])
    return snap


@bp.route(PREFIX+'/catalog-scans/<id_>')
@api
def scan_status(c,u,id_):
    snap=accessible_scan(c,u,id_)
    snap.pop('items_json');snap.pop('scope_json')
    return jsonify(snap)


@bp.route(PREFIX+'/catalog-snapshots/<id_>/products')
@api
def catalog_items(c,u,id_):
    snap=accessible_scan(c,u,id_)
    items=loads(snap['items_json'],[])
    q=request.args.get('search','').strip().casefold()
    filtered=[i for i in items if q in (i['name']+' '+i['sku_code']).casefold()]
    page=max(1,int(request.args.get('page',1)))
    size=min(100,max(1,int(request.args.get('per_page',50))))
    return jsonify(items=filtered[(page-1)*size:page*size],total=len(filtered),catalog_total=len(items),complete=bool(snap['complete']),page=page)


@bp.route(PREFIX+'/plans',methods=['POST'])
@api
def plans(c,u):
    return jsonify(id=create_plan(c,u,request.get_json())),202


@bp.route(PREFIX+'/plans/<id_>')
@api
def plan_status(c,u,id_):
    p,req=accessible_plan(c,u,id_)
    p.pop('request_json')
    p['summary']=loads(p.pop('summary_json'))
    p['scope']={k:req[k] for k in ('operation','source_site_id','target_sites','reason')}
    p['items']=[{'id':i['id'],**public_detail(u,loads(i['detail_json']))} for i in rows(c,'SELECT * FROM stock_sync_plan_items WHERE plan_id=? ORDER BY site_id,resource_key',(id_,))]
    return jsonify(p)


@bp.route(PREFIX+'/jobs',methods=['POST'])
@api
def jobs(c,u):
    return jsonify(id=confirm(c,u,request.get_json())),202


@bp.route(PREFIX+'/jobs')
@api
def history(c,u):
    page=max(1,int(request.args.get('page',1)))
    raw=rows(c,'SELECT * FROM stock_sync_jobs '+('' if u['superadmin'] else 'WHERE actor_id=? ')+
        'ORDER BY created_at DESC LIMIT 50 OFFSET ?',([u['id']] if not u['superadmin'] else [])+[(page-1)*50])
    result=[]
    for j in raw:
        try:
            accessible_job(c,u,j['id'])
            result.append({k:j[k] for k in ('id','status','created_at','completed_at')})
        except SyncError: pass
    return jsonify(items=result,page=page)


@bp.route(PREFIX+'/jobs/<id_>')
@api
def job_status(c,u,id_):
    j,p,req=accessible_job(c,u,id_)
    items=[]
    counts={}
    for item in rows(c,'SELECT * FROM stock_sync_job_items WHERE job_id=? ORDER BY id',(id_,)):
        detail=loads(one(c,'SELECT detail_json FROM stock_sync_plan_items WHERE id=?',(item['plan_item_id'],))['detail_json'])
        counts[item['status']]=counts.get(item['status'],0)+1
        items.append({'id':item['id'],'status':item['status'],'error':item['error'],'attempts':item['attempts'],
            'controls_saved':bool(item['intent_json']),'after':loads(item['after_json']) if item['after_json'] else None,
            'detail':public_detail(u,detail)})
    return jsonify(id=j['id'],status=j['status'],counts=counts,items=items,scope=req['target_sites'],created_at=j['created_at'])


@bp.route(PREFIX+'/jobs/<id_>/cancel',methods=['POST'])
@api
def cancel_job(c,u,id_):
    j,_,_=accessible_job(c,u,id_)
    return jsonify(status=cancel(c,u,j))


@bp.route(PREFIX+'/jobs/<id_>/retry-plan',methods=['POST'])
@api
def retry(c,u,id_):
    j,p,req=accessible_job(c,u,id_)
    ids=request.get_json().get('item_ids')
    all_items=rows(c,'SELECT * FROM stock_sync_job_items WHERE job_id=?',(id_,))
    if not isinstance(ids,list) or not ids: raise SyncError('INVALID_SELECTION',status=400)
    items=[x for x in all_items if x['id'] in ids]
    if len(items)!=len(set(ids)) or any(i['status'] not in RETRYABLE for i in items):
        raise SyncError('NOT_RETRYABLE','仅失败、取消或冲突项目可重试；待核实项目需先核实')
    details=[loads(one(c,'SELECT detail_json FROM stock_sync_plan_items WHERE id=?',(i['plan_item_id'],))['detail_json']) for i in items]
    req['target_site_ids']=sorted({d['site_id'] for d in details})
    req['target_sites']=[s for s in req['target_sites'] if s['id'] in req['target_site_ids']]
    req['retry_map_ids']=[d['map_id'] for d in details]
    req['retry_control_ids']=[f'{i["id"]}:{n}' for i,d in zip(items,details) for n,_ in enumerate(d.get('changes',[]))]
    req['retry_of']=id_
    new=uid()
    c.execute("INSERT INTO stock_sync_plans(id,actor_id,request_json,status,created_at,expires_at) VALUES(?,?,?,'building',?,?)",(new,u['id'],dumps(req),stamp(),stamp(600)))
    enqueue(c,'plan',new);event(c,'retry_preview',u['id'],new,{'job_id':id_,'item_ids':ids});c.commit()
    return jsonify(id=new),202


@bp.route(PREFIX+'/controls')
@api
def controls(c,u):
    ids=[s['id'] for s in target_sites(c,u,{'mode':'all_authorized'})]
    if not ids: return jsonify(items=[])
    result=rows(c,'''SELECT c.*,k.sku_code,k.name AS sku_name FROM stock_sync_controls c
        JOIN inv_skus k ON k.id=c.sku_id WHERE c.active=1 AND c.site_id IN ('''+','.join('?' for _ in ids)+') ORDER BY c.created_at DESC',ids)
    for r in result: r.pop('source_json',None)
    return jsonify(items=result)
