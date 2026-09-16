"""Immutable plans, explicit intersections and input fingerprints."""
from copy import deepcopy

from stock_sync_common import SyncError, uid, rows, one, dumps, loads, digest, stamp, now, parse_time, event
from stock_sync_permissions import actor, target_sites, reference_site, visible_site, can_release, require_object
from stock_sync_catalog import enqueue, select_items
from stock_sync_supply import mappings, mapping_hash, snapshot, strategy_conflict
from stock_sync_policy import resolve_target_stock, matches
from stock_sync_woo import Woo, physical_scope, stock_state, site_identity, managed_resource_keys

OPERATIONS = {'manual_hold','release_hold','quantity','reference_status'}


def create_plan(c,u,data):
    data = deepcopy(data)
    if data.get('operation') not in OPERATIONS:
        raise SyncError('INVALID_OPERATION',status=400)
    if not str(data.get('reason','')).strip():
        raise SyncError('REASON_REQUIRED','请填写操作原因',400)
    if len(str(data['reason'])) > 500:
        raise SyncError('INVALID_INPUT','原因最多 500 字',400)
    targets = target_sites(c,u,data.get('target_scope') or {})
    source_id = data.get('source_site_id') if data['operation']=='reference_status' else None
    if source_id:
        source_id = int(source_id)
        reference_site(c,u,source_id)
        targets = [s for s in targets if s['id'] != source_id]
    if data['operation']=='reference_status' and not source_id:
        raise SyncError('SOURCE_REQUIRED',status=400)
    if not targets:
        raise SyncError('EMPTY_TARGET','请至少选择一个目标站点',400)
    snap = one(c,'SELECT * FROM stock_sync_catalog_snapshots WHERE id=?',(data.get('catalog_snapshot_id'),))
    if not snap or (snap['actor_id'] != u['id'] and not u['superadmin']):
        raise SyncError('FORBIDDEN','无权使用该目录',403)
    if parse_time(snap['expires_at']) < now():
        raise SyncError('CATALOG_EXPIRED','目录过期，请重新扫描')
    if snap['site_id'] != source_id:
        raise SyncError('SOURCE_MISMATCH','目录与操作来源不一致',400)
    if not source_id:
        target_sites(c,u,loads(snap['scope_json']))
    chosen = select_items(snap,data.get('selection') or {})
    if data['operation']=='release_hold':
        ids = data.get('control_ids')
        if not isinstance(ids,list) or not ids or any(not isinstance(x,str) for x in ids):
            raise SyncError('CONTROL_IDS_REQUIRED','请选择具体停售控制',400)
        allowed_maps = {m for i in chosen for m in i['map_ids']}
        for id_ in ids:
            control = one(c,'SELECT * FROM stock_sync_controls WHERE id=? AND active=1',(id_,))
            if not control or control['kind']!='manual_hold' or control['map_id'] not in allowed_maps or control['site_id'] not in {s['id'] for s in targets} or not can_release(u,control):
                raise SyncError('CONTROL_FORBIDDEN','不能解除该控制或控制不在已选范围',403)
    # Keep only recognized options. Client supplied payloads/URLs/permissions are ignored.
    request = {'operation':data['operation'],'source_site_id':source_id,'catalog_snapshot_id':snap['id'],
        'target_site_ids':[s['id'] for s in targets],'target_sites':[visible_site(s) for s in targets],
        'selected_items':chosen,'reason':str(data['reason']).strip(),'selection':data.get('selection'),
        'control_ids':data.get('control_ids',[]) if data['operation']=='release_hold' else [],
        'confirm_available':data.get('confirm_available') is True and data['operation']=='release_hold',
        'reference_sync_mode':data.get('reference_sync_mode','status_both')}
    if request['reference_sync_mode'] not in ('status_both','sold_out_only'):
        raise SyncError('INVALID_INPUT',status=400)
    id_ = uid()
    c.execute('''INSERT INTO stock_sync_plans(id,actor_id,request_json,status,created_at,expires_at)
        VALUES(?,?,?,'building',?,?)''',(id_,u['id'],dumps(request),stamp(),stamp(600)))
    enqueue(c,'plan',id_)
    event(c,'plan_requested',u['id'],id_,{'operation':request['operation'],'targets':request['target_site_ids']})
    c.commit()
    return id_


def source_evidence(c,u,request,woo,sku_id):
    site = reference_site(c,u,request['source_site_id'])
    chosen = [i for i in request['selected_items'] if i.get('sku_id') == sku_id]
    evidence = []
    for item in chosen:
        if len(item['map_ids']) != 1:
            raise SyncError('AMBIGUOUS_MAPPING')
        m = one(c,'SELECT * FROM inv_site_sku_map WHERE id=? AND is_active=1',(item['map_ids'][0],))
        if not m or mapping_hash(m) != item['mapping_hashes'].get(str(m['id'])):
            raise SyncError('MAPPING_CHANGED')
        if m['site_id'] != site['id']:
            raise SyncError('MAPPING_CHANGED')
        c.commit()
        remote = woo.read(site,m)
        woo.validate(remote,m)
        supply = snapshot(c,m,site)
        evidence.append({'map_id':m['id'],'mapping_hash':mapping_hash(m),'site_identity':site_identity(site),'state':stock_state(remote),
            'qty_per_item':m['qty_per_item'],'pools':sorted({s['pool_id'] for s in supply['sources']})})
    if not evidence:
        raise SyncError('UNMAPPED')
    if len({e['state']['stock_status'] for e in evidence}) != 1:
        raise SyncError('SOURCE_STATUS_CONFLICT')
    if evidence[0]['state']['stock_status']=='onbackorder':
        raise SyncError('BACKORDER_POLICY_CONFLICT')
    return evidence


def evaluate(c,u,request,m,woo,*,source_cache=None):
    woo.connection = c
    sites = target_sites(c,u,{'mode':'explicit_sites','site_ids':request['target_site_ids']})
    site = next((s for s in sites if s['id']==m['site_id']),None)
    if not site:
        raise SyncError('PERMISSION_REVOKED',status=403)
    current_maps = mappings(c,[site['id']])
    alias = [x for x in current_maps if x['wc_product_id']==m['wc_product_id'] and (x['wc_variation_id'] or 0)==(m['wc_variation_id'] or 0)]
    if len(alias)!=1 or alias[0]['id']!=m['id']:
        raise SyncError('AMBIGUOUS_MAPPING')
    scope = physical_scope(c,site,m,request['target_site_ids'],request.get('source_site_id'))
    mh = mapping_hash(m)
    old_binding = one(c,'SELECT * FROM stock_sync_bindings WHERE map_id=?',(m['id'],))
    bound_baseline=loads(old_binding['baseline_json']) if old_binding else {}
    if old_binding and (old_binding['mapping_hash'] != mh or bound_baseline.get('resource_key',scope['resource_key'])!=scope['resource_key']):
        raise SyncError('MAPPING_CHANGED','映射已改变，需超级管理员重新绑定')
    for other in rows(c,'SELECT map_id,baseline_json FROM stock_sync_bindings WHERE map_id<>?',(m['id'],)):
        if scope['resource_key'] in loads(other['baseline_json']).get('managed_resource_keys',[]):
            raise SyncError('MAPPING_CHANGED','该物理资源仍属于历史绑定，不能通过新建映射绕过控制')
    controls = rows(c,'SELECT * FROM stock_sync_controls WHERE map_id=? AND active=1 ORDER BY created_at,id',(m['id'],))
    if any(x['mapping_hash']!=mh for x in controls):
        raise SyncError('MAPPING_CHANGED','旧控制不能自动作用到新映射')
    c.commit()
    remote = woo.read(site,m)
    woo.validate(remote,m)
    supply = snapshot(c,m,site)
    baseline = loads(old_binding['baseline_json']) if old_binding else {
        'mode':'quantity' if remote['manage_stock'] is True else 'status','state':stock_state(remote),'resource_key':scope['resource_key'],
        'managed_resource_keys':managed_resource_keys(c,site,m)}
    if request['operation']=='quantity' and strategy_conflict(supply):
        raise SyncError('PUBLISH_STRATEGY_CONFLICT','共享库存池仍有有效配额策略，请先核对并停用旧发布')
    changes, proposed = [], deepcopy(controls)
    if request['operation']=='manual_hold':
        if not any(x['id'] in request.get('retry_control_ids',[]) and x['kind']=='manual_hold' for x in controls):
            changes.append({'kind':'manual_hold','stock_status':'outofstock','protection':'superadmin' if u['superadmin'] else 'owner'})
    elif request['operation']=='release_hold':
        remove = [x for x in controls if x['id'] in request['control_ids']]
        if any(not can_release(u,x) for x in remove):
            raise SyncError('CONTROL_FORBIDDEN',status=403)
        proposed = [x for x in proposed if x['id'] not in request['control_ids']]
        changes.extend({'release':x['id']} for x in remove)
        if request['confirm_available'] and supply['result']['mode']=='status' and baseline['mode']=='status':
            proposed = [x for x in proposed if x['kind']!='manual_availability']
            changes.append({'kind':'manual_availability','stock_status':'instock','protection':'owner'})
    source = []
    if request['operation']=='reference_status':
        if source_cache is not None and m['sku_id'] in source_cache:
            source = source_cache[m['sku_id']]
            if isinstance(source,SyncError):
                raise source
        else:
            source = source_evidence(c,u,request,woo,m['sku_id'])
            if source_cache is not None:
                source_cache[m['sku_id']] = source
        if any(e['qty_per_item']!=m['qty_per_item'] for e in source):
            raise SyncError('UNIT_NOT_COMPARABLE')
        pools = sorted({s['pool_id'] for s in supply['sources']})
        if not pools or any(e['pools']!=pools for e in source):
            raise SyncError('SUPPLY_SCOPE_CONFLICT','来源与目标的供货范围不同或未明确')
        status = source[0]['state']['stock_status']
        if request['reference_sync_mode']=='sold_out_only' and status=='instock':
            raise SyncError('SOLD_OUT_ONLY_SKIP')
        proposed = [x for x in proposed if x['kind']!='reference_snapshot']
        changes.append({'kind':'reference_snapshot','stock_status':status,'protection':'reference',
            'source':{'site_id':request['source_site_id'],'evidence':source}})
    proposed.extend(x for x in changes if 'release' not in x)
    intended, reason = resolve_target_stock(baseline,supply['result'],proposed)
    if intended.get('manage_stock') is True and strategy_conflict(supply):
        raise SyncError('PUBLISH_STRATEGY_CONFLICT')
    inputs = {'mapping_hash':mh,'controls':controls,'binding':old_binding,'scope':scope,
        'before':stock_state(remote),'parent':remote.get('_parent_manage_stock'),
        'supply':supply,'source':source,'owner':site.get('manager')}
    return {'map_id':m['id'],'site_id':site['id'],'sku_id':m['sku_id'],'name':m['sku_name'],
        'sku_code':m['sku_code'],'qty_per_item':m['qty_per_item'],'mapping':m,'scope':scope,
        'resource_key':scope['resource_key'],'before':stock_state(remote),'intended':intended,
        'baseline':baseline,'changes':changes,'controls':controls,'supply':supply,'source':source,
        'inputs_hash':digest(inputs),'reason':reason,'decision':'unchanged' if matches(remote,intended) else 'change',
        'mode_changed':remote['manage_stock']!=intended['manage_stock'],
        'backorders_changed':remote['backorders']!=intended['backorders'],
        'restores_sales':remote['stock_status']=='outofstock' and intended['stock_status']=='instock'}


def build_plan(c,id_,woo=None):
    woo = woo or Woo()
    plan = one(c,'SELECT * FROM stock_sync_plans WHERE id=?',(id_,))
    request = loads(plan['request_json'])
    summary = {'change':0,'unchanged':0,'skipped':0,'conflict':0,'unmatched_pairs':0}
    try:
        u = actor(c,plan['actor_id'])
        require_object(c,u,plan,request['target_site_ids'],request.get('source_site_id'))
        candidates = mappings(c,request['target_site_ids'])
        selected = request['selected_items']
        if request['operation']=='reference_status':
            skus = {i['sku_id'] for i in selected if i.get('sku_id')}
            candidates = [m for m in candidates if m['sku_id'] in skus]
            missing = [(i,sid) for i in selected for sid in request['target_site_ids']
                if not i.get('sku_id') or not any(m['site_id']==sid and m['sku_id']==i['sku_id'] for m in candidates)]
            for i,sid in missing:
                detail = {'site_id':sid,'name':i['name'],'decision':'skipped','reason':'UNMAPPED','catalog_item_id':i['id']}
                c.execute('INSERT INTO stock_sync_plan_items VALUES(?,?,?,?,?,?,?)',(uid(),id_,f'missing:{sid}:{i["id"]}',None,sid,'skipped',dumps(detail)))
                summary['skipped']+=1; summary['unmatched_pairs']+=1
        else:
            selected_maps = {mid: i['mapping_hashes'][str(mid)] for i in selected for mid in i['map_ids']}
            candidates = [m for m in candidates if m['id'] in selected_maps]
            if any(mapping_hash(m)!=selected_maps[m['id']] for m in candidates):
                raise SyncError('MAPPING_CHANGED')
            eligible = {mid for mid in selected_maps if next(i for i in selected if mid in i['map_ids'])['site_id'] in request['target_site_ids']}
            if eligible != {m['id'] for m in candidates}:
                raise SyncError('MAPPING_CHANGED')
        seen, source_cache = {}, {}
        if 'retry_map_ids' in request:
            candidates = [m for m in candidates if m['id'] in request['retry_map_ids']]
            if {m['id'] for m in candidates} != set(request['retry_map_ids']):
                raise SyncError('MAPPING_CHANGED')
        for m in candidates:
            u = actor(c,plan['actor_id'])
            try:
                detail = evaluate(c,u,request,m,woo,source_cache=source_cache)
            except SyncError as exc:
                detail = {'site_id':m['site_id'],'map_id':m['id'],'name':m['sku_name'],'sku_code':m['sku_code'],
                    'decision':'skipped' if exc.code=='SOLD_OUT_ONLY_SKIP' else 'conflict','reason':exc.code,
                    'resource_key':f'conflict:{m["id"]}'}
            key = detail['resource_key']
            if key in seen:
                # Ambiguous aliases are never allowed to silently discard one map's
                # controls. Both items are conflicts until the duplicate is resolved.
                old_id, old_decision = seen[key]
                old = one(c,'SELECT detail_json FROM stock_sync_plan_items WHERE id=?',(old_id,))
                prior = loads(old['detail_json']); prior.update(decision='conflict',reason='PHYSICAL_SCOPE_CONFLICT')
                c.execute("UPDATE stock_sync_plan_items SET decision='conflict',detail_json=? WHERE id=?",(dumps(prior),old_id))
                if old_decision!='conflict':
                    summary[old_decision]-=1;summary['conflict']+=1
                detail.update(decision='conflict',reason='PHYSICAL_SCOPE_CONFLICT')
                key = f'duplicate:{m["id"]}'
            item_id = uid()
            c.execute('INSERT INTO stock_sync_plan_items VALUES(?,?,?,?,?,?,?)',(item_id,id_,key,m['id'],m['site_id'],detail['decision'],dumps(detail)))
            seen[key] = (item_id,detail['decision'])
            summary[detail['decision']]+=1
            c.commit()
        summary.update(target_sites=len(request['target_site_ids']),selected_leaves=len(selected),
            selected_products=len({(i['site_id'],i['product_id']) for i in selected}),
            selected_skus=len({i.get('sku_id') for i in selected if i.get('sku_id')}),
            target_resources=len(candidates))
        c.execute("UPDATE stock_sync_plans SET status='ready',summary_json=?,expires_at=? WHERE id=?",(dumps(summary),stamp(600),id_))
    except SyncError as exc:
        c.execute("UPDATE stock_sync_plans SET status='failed',error=? WHERE id=?",(exc.code,id_))
    except Exception:
        c.rollback()
        c.execute("UPDATE stock_sync_plans SET status='failed',error='PLAN_BUILD_FAILED' WHERE id=?",(id_,))
        raise
    finally:
        c.commit()
