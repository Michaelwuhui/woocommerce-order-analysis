/* Explicit server snapshots and immutable confirmations. No browser-side writes to Woo. */
(() => {
  'use strict';
  const $ = id => document.getElementById(`ss-${id}`);
  const base = '/api/product-manager/stock-sync';
  let csrf = document.getElementById('stock-sync').dataset.csrf;
  let options, snapshot, complete = false, catalog = [], selected = new Set(), targets = new Set();
  let page = 1, plan, job, generation = 0, pollGeneration = 0, selectionMode = 'explicit', frozenFilter = '';
  const labels = {change:'待修改',unchanged:'原已一致',skipped:'跳过',conflict:'需重新预览',failed:'失败',pending:'待执行',running:'执行中',verified_success:'回读已确认',uncertain:'结果待核实',superseded:'已被新控制替代',cancelled:'已取消',queued:'已排队',succeeded:'全部已核实',partial_failed:'部分完成',requires_review:'需要核实',cancel_requested:'正在取消后续项目',building:'正在生成预览'};
  const reasons = {UNMAPPED:'没有有效映射',INVENTORY_UNKNOWN:'库存依据未知',PHYSICAL_SCOPE_CONFLICT:'实际影响范围不能隔离',PARENT_STOCK_SHARED:'口味共享父级库存',UNIT_NOT_COMPARABLE:'销售单位不可直接比较',SOURCE_STATUS_CONFLICT:'参照别名状态冲突',SOURCE_INCOMPLETE:'目录未完整读取',PUBLISH_STRATEGY_CONFLICT:'旧配额发布策略冲突',BACKORDER_POLICY_CONFLICT:'缺货预订策略冲突',ORDER_SYNC_BACKLOG:'订单同步存在积压或失败',MAPPING_CHANGED:'映射已改变',REMOTE_STATE_CHANGED:'预览后的状态已改变',MANUAL_HOLD_ACTIVE:'人工停售保护有效',REFERENCE_HOLD_ACTIVE:'参照停售有效',ACTUAL_STOCK:'按当前可售数量',ACTUAL_STOCK_ZERO:'当前实际可售为零',STATUS_SOURCE:'按可靠供货状态',READBACK_MISMATCH:'目标站回读不一致',RESOURCE_BUSY:'资源正在执行或等待核实',EXTERNAL_STOCK_STALE:'外部库存数据过期',SUPPLY_SCOPE_CONFLICT:'参照与目标供货来源不同',PERMISSION_REVOKED:'执行权限已撤销',WRITE_RESULT_UNKNOWN:'网站结果尚未确认',PLAN_STALE:'预览已过期'};
  const el = (tag, text, cls) => { const n=document.createElement(tag); if(text!==undefined)n.textContent=text; if(cls)n.className=cls; return n; };
  const message = (text,bad=false) => { $('message').textContent=text; $('message').classList.toggle('ss-bad',bad); };
  const error = e => message(e.message || String(e),true);
  async function api(path,method='GET',body) {
    const r=await fetch(base+path,{method,headers:{'Content-Type':'application/json','X-CSRF-Token':csrf},body:body===undefined?undefined:JSON.stringify(body)});
    let data;try{data=await r.json();}catch{throw new Error('登录已失效或服务未返回 JSON，请刷新页面。');}
    if(!r.ok)throw new Error(`${data.error || '请求失败'}${data.code?' ('+data.code+')':''}`);
    return data;
  }
  const sleep = ms => new Promise(resolve=>setTimeout(resolve,ms));
  function invalidate(){ generation++; plan=null;$('execute').disabled=true;$('plan').replaceChildren();$('plan-summary').textContent='选择已改变，请重新生成差异预览。';$('plan-scope').textContent=''; }
  function sourceChange(){invalidate();$('scan-info').textContent='操作或来源已改变，请重新读取目录。';snapshot=null;complete=false;catalog=[];selected.clear();renderProducts();const ref=$('operation').value==='reference_status';$('source-wrap').classList.toggle('ss-hidden',!ref);$('available-wrap').classList.toggle('ss-hidden',$('operation').value!=='release_hold');$('controls-wrap').classList.toggle('ss-hidden',$('operation').value!=='release_hold');if(ref)targets.delete(Number($('source').value));renderSites();loadControls().catch(error);}
  function option(select,value,text){const o=el('option',text);o.value=value;select.append(o);}
  function visibleSites(){return options.target_sites.filter(s=>(!$('manager').value||s.manager===$('manager').value)&&(!$('country').value||s.country===$('country').value));}
  function renderSites(){
    $('sites').replaceChildren();
    for(const s of visibleSites()){
      const label=el('label'),check=el('input');check.type='checkbox';check.checked=targets.has(s.id);check.disabled=$('operation').value==='reference_status'&&s.id===Number($('source').value);
      check.onchange=()=>{check.checked?targets.add(s.id):targets.delete(s.id);invalidate();};
      label.append(check,el('span',`${s.url} · ${s.manager||'无负责人'} · ${s.country||''}${!s.available?'（凭据不完整或禁用）':''}`));$('sites').append(label);
    }
  }
  function filtered(){const q=$('search').value.trim().toLocaleLowerCase();return catalog.filter(i=>(i.name+' '+i.sku_code).toLocaleLowerCase().includes(q));}
  function renderProducts(){
    $('products').replaceChildren();const items=filtered();const per=30;page=Math.max(1,Math.min(page,Math.ceil(items.length/per)||1));
    const groups=new Map();for(const i of items.slice((page-1)*per,page*per)){const k=`${i.site_id}:${i.product_id}`;if(!groups.has(k))groups.set(k,[]);groups.get(k).push(i);}
    for(const [key,list] of groups){
      const group=el('div',undefined,'ss-group'),head=el('label'),check=el('input');check.type='checkbox';
      const full=catalog.filter(i=>`${i.site_id}:${i.product_id}`===key),count=full.filter(i=>selected.has(i.id)).length;check.checked=count===full.length;check.indeterminate=count>0&&count<full.length;
      check.onchange=()=>{for(const i of full){check.checked?selected.add(i.id):selected.delete(i.id);}invalidate();renderProducts();};
      head.append(check,el('strong',` ${list[0].name.split(' / ')[0]} · ${full.length} 个商品/口味`));group.append(head);
      const leaves=el('div',undefined,'ss-leaves');for(const i of list){const label=el('label'),box=el('input');box.type='checkbox';box.checked=selected.has(i.id);box.onchange=()=>{box.checked?selected.add(i.id):selected.delete(i.id);invalidate();renderProducts();};label.append(box,el('span',`${i.name} · ${i.sku_code||'未映射'} · 站点 ${i.site_id}${i.qty_per_item?' · 每件 '+i.qty_per_item+' 基础单位':''}`));leaves.append(label);}group.append(leaves);$('products').append(group);
    }
    const chosen=catalog.filter(i=>selected.has(i.id));$('selection').textContent=`已选 ${new Set(chosen.map(i=>`${i.site_id}:${i.product_id}`)).size} 款 · ${chosen.length} 个商品/口味 · ${new Set(chosen.filter(i=>i.sku_id).map(i=>i.sku_id)).size} 个有效 SKU；目录 ${catalog.length} 项${complete?'（完整）':'（尚不完整）'}`;
    $('page').textContent=`第 ${page} / ${Math.ceil(items.length/per)||1} 页，当前筛选 ${items.length} 项`;$('prev').disabled=page===1;$('next').disabled=page*per>=items.length;$('all').disabled=!complete;$('filtered').disabled=!complete;
  }
  async function scan(){
    invalidate();const g=generation;$('scan').disabled=true;catalog=[];selected.clear();snapshot=null;complete=false;renderProducts();$('scan-info').textContent='正在读取新目录，尚未确认完整性…';
    try{
      const body=$('operation').value==='reference_status'?{source_site_id:Number($('source').value)}:{target_scope:{mode:'all_authorized'}};
      const created=await api('/catalog-scans','POST',body);snapshot=created.id;
      for(;;){if(g!==generation)return;const s=await api('/catalog-scans/'+snapshot);$('scan-info').textContent=`目录${s.complete?'已完整读取':'读取中'}：${s.progress} 个商品/口味${s.error?'；'+s.error:''}`;if(['complete','incomplete','failed'].includes(s.status)){complete=!!s.complete;break;}await sleep(1000);}
      let p=1;for(;;){const r=await api(`/catalog-snapshots/${snapshot}/products?page=${p}&per_page=100`);if(g!==generation)return;catalog.push(...r.items);if(catalog.length>=r.total)break;p++;}
      renderProducts();message(complete?'目录读取完成，可跨分页全选或展开口味。':'目录读取不完整；只能选择已完整读取的明确项目，不能全选。',!complete);
    }finally{$('scan').disabled=false;}
  }
  function selection(){
    if(selectionMode==='all')return {mode:'all',excluded_catalog_item_ids:catalog.filter(i=>!selected.has(i.id)).map(i=>i.id)};
    if(selectionMode==='filtered_all'){
      const pool=catalog.filter(i=>(i.name+' '+i.sku_code).toLocaleLowerCase().includes(frozenFilter.toLocaleLowerCase()));
      // If later manual selection expanded beyond that filter, freeze exact IDs.
      if([...selected].every(id=>pool.some(i=>i.id===id)))return {mode:'filtered_all',filter:{search:frozenFilter},excluded_catalog_item_ids:pool.filter(i=>!selected.has(i.id)).map(i=>i.id)};
    }
    return {mode:'explicit',catalog_item_ids:[...selected]};
  }
  const checkedControls=()=>[...$('controls').querySelectorAll('input:checked')].map(n=>n.value);
  async function preview(){
    if(!snapshot||!selected.size||!targets.size)throw new Error('请读取目录，并选择商品和目标站点。');
    invalidate();const g=generation;$('preview').disabled=true;
    try{const body={operation:$('operation').value,source_site_id:$('operation').value==='reference_status'?Number($('source').value):null,catalog_snapshot_id:snapshot,selection:selection(),target_scope:{mode:'explicit_sites',site_ids:[...targets]},reason:$('reason').value,control_ids:checkedControls(),confirm_available:$('available').checked};const r=await api('/plans','POST',body);await waitPlan(r.id,g);}finally{$('preview').disabled=false;}
  }
  function table(headers){const t=el('table',undefined,'ss-table'),head=el('tr');for(const h of headers)head.append(el('th',h));const th=el('thead');th.append(head);t.append(th);const body=el('tbody');t.append(body);return [t,body];}
  function state(s){if(!s)return '—';return `${s.stock_status==='instock'?'有货':s.stock_status==='outofstock'?'售完':s.stock_status||'未知'} / ${s.manage_stock===true?'数量 '+(s.stock_quantity??'未知'):'状态管理'} / 预订 ${s.backorders||'未知'}`;}
  async function waitPlan(id,g){
    for(;;){const p=await api('/plans/'+id);if(g!==generation)return;$('plan-summary').textContent=labels[p.status]||p.status;if(p.status==='ready'){plan=p;renderPlan();return;}if(p.status==='failed')throw new Error(reasons[p.error]||p.error);await sleep(1000);}
  }
  function renderPlan(){
    const s=plan.summary;$('plan-summary').textContent=`预计修改 ${s.change} 个商品/变体，涉及 ${s.target_sites} 个站点；原已一致 ${s.unchanged}，跳过 ${s.skipped}，冲突 ${s.conflict}。预览有效至 ${new Date(plan.expires_at).toLocaleTimeString()}`;
    $('plan-scope').textContent=plan.scope.target_sites.map(s=>`${s.url}（${s.manager||'无负责人'}）`).join('；');
    const [t,b]=table(['执行','站点 / 商品 / SKU','来源与控制','当前状态','预计状态','处理结论']);
    for(const i of plan.items){const tr=el('tr'),select=el('td'),box=el('input');box.type='checkbox';box.value=i.id;box.checked=['change','unchanged'].includes(i.decision);box.disabled=!box.checked;select.append(box);tr.append(select,el('td',`${i.site_id} / ${i.name} / ${i.sku_code||''}${i.qty_per_item?' · '+i.qty_per_item+' 单位':''}`),el('td',i.source?.length?`参照站 ${plan.scope.source_site_id}`:(i.controls?.length?i.controls.map(c=>`${c.kind}：${c.reason}`).join('；'):'当前供货来源')),el('td',state(i.before)),el('td',state(i.intended)),el('td',`${labels[i.decision]} · ${reasons[i.reason]||i.reason}${i.mode_changed?'；库存管理模式改变':''}${i.backorders_changed?'；预订规则改变':''}${i.restores_sales?'；将恢复销售':''}`,i.decision==='conflict'?'ss-bad':''));b.append(tr);}
    $('plan').replaceChildren(t);$('execute').disabled=!plan.items.some(i=>['change','unchanged'].includes(i.decision));message('请核对具体站点、商品及库存管理方式后确认执行。');
  }
  async function execute(){
    if(!plan)throw new Error('请重新预览。');const ids=[...$('plan').querySelectorAll('input:checked')].map(n=>n.value);if(!ids.length)throw new Error('没有选中可执行项目。');$('execute').disabled=true;
    try{const key=plan.idempotency_key||(plan.idempotency_key=crypto.randomUUID());const r=await api('/jobs','POST',{plan_id:plan.id,plan_version:plan.version,idempotency_key:key,accepted_item_ids:ids});job=r.id;await watchJob(job);}catch(e){$('execute').disabled=false;throw e;}
  }
  async function watchJob(id){const pg=++pollGeneration;job=id;for(;;){const j=await api('/jobs/'+id);if(pg!==pollGeneration)return;renderJob(j);if(!['queued','running','cancel_requested'].includes(j.status)){message(`任务${labels[j.status]||j.status}，请查看逐项回读结果。`,j.status!=='succeeded');await loadHistory();await loadControls();return;}await sleep(1000);}}
  function renderJob(j){
    $('job-summary').textContent=`任务 ${j.id} · ${labels[j.status]||j.status} · ${Object.entries(j.counts).map(([k,v])=>`${labels[k]||k} ${v}`).join('，')}`;
    const [t,b]=table(['重试','站点 / 商品','结果','回读状态','控制与异常']);
    for(const i of j.items){const tr=el('tr'),cell=el('td'),box=el('input');box.type='checkbox';box.value=i.id;box.checked=['failed','conflict','cancelled','superseded'].includes(i.status);box.disabled=!box.checked;cell.append(box);tr.append(cell,el('td',`${i.detail.site_id} / ${i.detail.name}`),el('td',labels[i.status]||i.status),el('td',state(i.after)),el('td',`${i.controls_saved?'控制意图已保存；':''}${i.error?(reasons[i.error]||i.error):''}${i.controls_saved&&!['verified_success','unchanged'].includes(i.status)?' 网站尚未确认同步':''}`));b.append(tr);}
    $('job').replaceChildren(t);$('cancel').disabled=!['queued','running','cancel_requested'].includes(j.status);$('retry').disabled=!j.items.some(i=>['failed','conflict','cancelled','superseded'].includes(i.status));
  }
  async function loadHistory(){const h=await api('/jobs');$('history').replaceChildren();for(const j of h.items){const b=el('button',`${new Date(j.created_at).toLocaleString()} · ${labels[j.status]||j.status} · ${j.id.slice(0,8)}`,'btn btn-sm btn-outline-secondary');b.onclick=()=>watchJob(j.id).catch(error);$('history').append(b);}}
  async function loadControls(){const r=await api('/controls');$('controls').replaceChildren();for(const c of r.items.filter(c=>c.kind==='manual_hold')){const label=el('label'),box=el('input');box.type='checkbox';box.value=c.id;box.disabled=c.protection==='superadmin'&&!options.superadmin;box.onchange=invalidate;label.append(box,el('span',`站点 ${c.site_id} / ${c.sku_name}：${c.reason}（${c.protection==='superadmin'?'超级管理员保护':'负责人控制'}）`));$('controls').append(label);}}
  const act=(id,fn)=>$(id).addEventListener('click',()=>Promise.resolve().then(fn).catch(error));
  $('operation').onchange=sourceChange;$('source').onchange=sourceChange;$('reason').oninput=invalidate;$('available').onchange=invalidate;
  $('manager').onchange=renderSites;$('country').onchange=renderSites;$('search').oninput=()=>{page=1;renderProducts();};
  act('scan',scan);act('all',()=>{selected=new Set(catalog.map(i=>i.id));selectionMode='all';invalidate();renderProducts();});act('filtered',()=>{selected=new Set(filtered().map(i=>i.id));selectionMode='filtered_all';frozenFilter=$('search').value.trim();invalidate();renderProducts();});act('clear',()=>{selected.clear();selectionMode='explicit';invalidate();renderProducts();});act('prev',()=>{page--;renderProducts();});act('next',()=>{page++;renderProducts();});
  act('select-sites',()=>{for(const s of visibleSites())if(!($('operation').value==='reference_status'&&s.id===Number($('source').value)))targets.add(s.id);invalidate();renderSites();});act('clear-sites',()=>{targets.clear();invalidate();renderSites();});act('preview',preview);act('execute',execute);
  act('cancel',async()=>{await api('/jobs/'+job+'/cancel','POST',{});await watchJob(job);});
  act('retry',async()=>{const ids=[...$('job').querySelectorAll('input:checked')].map(n=>n.value);invalidate();const g=generation,r=await api('/jobs/'+job+'/retry-plan','POST',{item_ids:ids});await waitPlan(r.id,g);$('plan-summary').scrollIntoView({behavior:'smooth'});});
  async function init(){options=await api('/options');csrf=options.csrf_token;for(const s of options.reference_sites)option($('source'),s.id,s.url);for(const m of new Set(options.target_sites.map(s=>s.manager).filter(Boolean)))option($('manager'),m,m);for(const c of new Set(options.target_sites.map(s=>s.country).filter(Boolean)))option($('country'),c,c);renderSites();renderProducts();await loadHistory();await loadControls();
    if(options.superadmin){$('admin').classList.remove('ss-hidden');for(const s of options.target_sites){const label=el('label'),box=el('input');box.type='checkbox';box.checked=options.reference_settings.some(r=>r.site_id===s.id&&r.enabled);box.onchange=async()=>{try{await api('/reference-sites/'+s.id,'PUT',{enabled:box.checked});message('共享参照设置已保存。');}catch(e){box.checked=!box.checked;error(e);}};label.append(box,el('span',s.url));$('reference-settings').append(label);}}
    message('请选择操作方式，读取本次商品目录。');}
  init().catch(error);
})();
