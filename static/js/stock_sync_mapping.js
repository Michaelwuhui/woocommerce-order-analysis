/* Explicit reviewed mappings, separate from inventory execution. */
(() => {
  'use strict';
  const $=id=>document.getElementById('ss-map-'+id), bridge=window.stockSyncMappingBridge;
  if(!bridge)return;
  const el=(tag,text)=>{const n=document.createElement(tag);if(text!==undefined)n.textContent=text;return n;};
  const scopeKey=()=>JSON.stringify(bridge.scope());
  const tell=(text,bad=false)=>{$('status').textContent=text;$('status').classList.toggle('ss-bad',bad);};
  const sleep=()=>new Promise(resolve=>setTimeout(resolve,1000));
  let data, savedScope, page=1, busy=false, initialized=false, choices=new Map();
  const states={mapped:'已有映射',conflict:'原记录冲突',unsupported:'不支持',unmatched:'信息不足',suggested:'自动匹配',new:'自动补建主档',align:'统一临时主档',review:'需人工核对',saved:'已保存',aligned:'已统一',unchanged:'原已一致',pending:'等待核实'};
  const editable=i=>i.writable&&['unmatched','suggested','review','new','align'].includes(i.state);
  const filtered=()=>{const q=$('search').value.trim().toLocaleLowerCase(),f=$('filter').value;return (data?.items||[]).filter(i=>(f==='all'||(f==='review'?['review','unmatched','conflict'].includes(i.state):f==='auto'?i.certain&&editable(i):!['mapped','unsupported'].includes(i.state)))&&(!q||(i.name+' '+i.site_url+' '+(data.skus.find(s=>s.id===choices.get(i.id)?.sku_id)?.sku_code||i.new_sku?.sku_code||'')).toLocaleLowerCase().includes(q)));};
  function selected(){return [...choices.entries()].filter(([,v])=>v.checked).map(([id,v])=>({id,sku_id:v.sku_id,qty_per_item:v.quantity,...(v.create_sku?{create_sku:true}:{})}));}
  function counts(){
    const picks=selected(),newCount=new Set(picks.filter(i=>i.create_sku).map(i=>data.items.find(x=>x.id===i.id)?.new_sku?.identity_key)).size;
    $('selection').textContent=`已自动选择或手选 ${picks.length} 项映射，需补建 ${newCount} 个统一产品主档；全部目录 ${data?.items.length||0} 项。可筛选查看自动匹配，未选项目保持原状。`;
    $('confirm').disabled=busy||!selected().length||!$('reviewed').checked;
  }
  function changed(){ $('reviewed').checked=false;counts(); }
  function render(){
    const pool=filtered(),size=20;page=Math.max(1,Math.min(page,Math.ceil(pool.length/size)||1));
    const table=el('table');table.className='ss-table';const head=el('tr');
    for(const title of ['选择','站点 / 商品 / 变体','识别依据','统一 SKU','每件折合数量'])head.append(el('th',title));
    const thead=el('thead');thead.append(head);table.append(thead);const body=el('tbody');table.append(body);
    for(const item of pool.slice((page-1)*size,page*size)){
      const choice=choices.get(item.id),tr=el('tr'),pick=el('input');pick.type='checkbox';pick.checked=choice.checked;pick.disabled=busy||!editable(item);pick.setAttribute('aria-label','选择映射 '+item.id);
      pick.onchange=()=>{choice.checked=pick.checked;changed();};const first=el('td');first.append(pick);tr.append(first);
      const productCell=el('td',`${item.site_url}\n${item.name}\n商品 ${item.product_id} / 变体 ${item.variation_id}`);
      if(item.wc_sku){const rawSku=el('div','站点 SKU：'+item.wc_sku);rawSku.className='ss-muted ss-map-chosen';productCell.append(rawSku);}tr.append(productCell);
      const evidence=el('td',`${states[item.state]||item.state}：${item.explanation||''}${item.writable?'':'；参照站只读，需由有权人员保存映射'}`);
      if(item.product_identity){const p=item.product_identity;const identity=el('div',`${p.brand} · ${p.model} · ${p.puff_count} 口 · ${p.flavor}`);identity.className='ss-muted ss-map-chosen';evidence.append(identity);}tr.append(evidence);
      const skuCell=el('td'),select=el('select');select.setAttribute('aria-label','映射 SKU '+item.id);
      const empty=el('option','请选择 SKU');empty.value='';select.append(empty);
      if(item.new_sku){const n=el('option','自动补建 · '+item.new_sku.name);n.value='__new';select.append(n);}
      const candidates=new Set(item.candidates||[]);
      for(const s of [...data.skus].sort((a,b)=>Number(candidates.has(b.id))-Number(candidates.has(a.id)))){
        const opt=el('option',`${candidates.has(s.id)?'建议 · ':''}${s.sku_code} · ${s.name}`);opt.value=s.id;select.append(opt);
      }
      select.value=choice.create_sku?'__new':choice.sku_id||'';select.disabled=busy||!editable(item)||item.state==='align';
      const chosenName=el('div');chosenName.className='ss-muted ss-map-chosen';
      const showChosen=()=>{const chosen=data.skus.find(s=>s.id===choice.sku_id);chosenName.textContent=choice.create_sku?`${item.new_sku.sku_code} · ${item.new_sku.warehouse_name}（供货状态管理）`:chosen?`${chosen.sku_code} · ${chosen.name}`:'';};showChosen();
      select.onchange=()=>{choice.create_sku=select.value==='__new';choice.sku_id=Number(select.value)||null;showChosen();changed();};skuCell.append(select,chosenName);tr.append(skuCell);
      const quantityCell=el('td'),quantity=el('input');quantity.type='number';quantity.min='1';quantity.max='100000';quantity.step='1';quantity.value=choice.quantity;quantity.style.width='95px';quantity.disabled=busy||!editable(item)||item.state==='align';quantity.setAttribute('aria-label','每件折合数量 '+item.id);
      quantity.oninput=()=>{choice.quantity=Number(quantity.value);changed();};quantityCell.append(quantity);tr.append(quantityCell);body.append(tr);
    }
    if(!pool.length){const tr=el('tr'),cell=el('td','当前筛选没有待处理商品，可切换“全部商品”查看已有映射。');cell.colSpan=5;tr.append(cell);body.append(tr);}
    $('table').replaceChildren(table);$('page').textContent=`第 ${page} / ${Math.ceil(pool.length/size)||1} 页，筛选 ${pool.length} 项`;
    $('prev').disabled=page===1;$('next').disabled=page*size>=pool.length;counts();
  }
  function accept(result){data=result;choices=new Map(data.items.map(i=>[i.id,{sku_id:i.proposed_sku_id,create_sku:i.state==='new',quantity:i.qty_per_item||1,checked:!!(i.certain&&editable(i))}]));page=1;$('filter').value=result.manual_required?'review':'pending';$('reviewed').checked=false;$('results').hidden=false;render();}
  async function readSuggestions(expected){
    const created=await bridge.api('/mapping-scans','POST',bridge.scope());
    for(;;){
      const result=await bridge.api('/mapping-scans/'+created.id);
      tell(`后台正在读取商品目录，已读取 ${result.progress} 项…`);
      if(result.status==='failed')throw new Error(result.error);
      if(result.complete){
        if(scopeKey()!==expected)throw new Error('站点范围已改变，请重新读取映射建议。');
        savedScope=expected;accept(result);return result;
      }
      await sleep();
    }
  }
  async function scan(){
    if(busy)return;if(!bridge.scope().target_scope.site_ids.length){tell('请先在第 3 步勾选目标站点。',true);return;}
    savedScope=scopeKey();busy=true;$('scan').disabled=true;$('scan').textContent='正在读取映射建议…';$('results').hidden=true;
    tell('正在提交映射目录读取请求…');
    try{
      const result=await readSuggestions(savedScope);
      tell(`已读取 ${result.items.length} 项，参考 ${result.identity_summary?.order_products||0} 条订单商品记录：${result.counts.mapped||0} 项已有映射，${result.auto_ready||0} 项已自动选择，${result.manual_required||0} 项需人工核对。自动选择包含 ${result.counts.new||0} 项待补建、${result.counts.align||0} 项待统一映射；确认保存后生效。`);
    }catch(e){tell(e.message,true);}finally{busy=false;$('scan').disabled=!bridge.allowed();$('scan').textContent='读取并生成映射建议';if(data)render();}
  }
  async function watchConfirmation(id, expected){
    for(;;){
      const result=await bridge.api('/mapping-confirmations/'+id);
      tell(`后台正在核实并保存映射：已核实 ${result.progress} / ${result.items.length} 项。关闭页面不影响已提交任务。`);
      if(result.status==='failed'){sessionStorage.removeItem('stock-sync-mapping-confirmation');throw new Error(result.error);}
      if(result.complete){
        sessionStorage.removeItem('stock-sync-mapping-confirmation');
        const n=(result.counts.saved||0)+(result.counts.aligned||0),m=result.counts.unchanged||0;
        if(data){const completed=new Map(result.items.map(i=>[i.id,i]));for(const item of data.items){if(completed.has(item.id)){item.state='mapped';item.explanation='本次已确认保存';choices.get(item.id).checked=false;}}}
        $('reviewed').checked=false;
        tell(`已保存 ${n} 条映射，${m} 条原已一致。正在更新库存同步目录…`);
        const refreshed=await bridge.refresh(expected);
        if(refreshed&&data&&result.items.some(i=>i.create_sku)&&data.items.some(editable)){
          savedScope=null;$('results').hidden=true;
          tell(`已保存 ${n} 条映射，正在按新主档更新剩余商品的映射建议…`);
          try{await readSuggestions(expected);}catch(e){throw new Error(`映射已保存，剩余建议刷新失败：${e.message}`);}
        }
        tell(`已保存 ${n} 条映射，${m} 条原已一致。${refreshed?'目录与剩余建议已更新，可继续处理或在第 4 步生成库存差异预览。':'当前站点范围已改变，请重新读取商品目录。'}`);
        return;
      }
      await sleep();
    }
  }
  async function confirm(){
    if(busy||!data)return;
    if(scopeKey()!==savedScope){tell('站点范围已改变，请重新读取映射建议。',true);return;}
    const items=selected();
    if(!items.length||!$('reviewed').checked){tell('请选择映射并确认已核对口味、规格和数量。',true);return;}
    if(items.some(i=>(!i.sku_id&&!i.create_sku)||!Number.isInteger(i.qty_per_item)||i.qty_per_item<1||i.qty_per_item>100000)){tell('所选项目必须选择已有 SKU 或自动补建，并填写有效的每件折合数量。',true);return;}
    if(!bridge.reason().trim()){tell('请先填写第 1 步的操作原因，再保存映射。',true);document.getElementById('ss-reason').focus();return;}
    busy=true;$('scan').disabled=true;render();tell('正在提交已确认的映射…');
    try{
      const r=await bridge.api('/mapping-scans/'+data.id+'/confirm','POST',{items,reviewed:true,reason:bridge.reason().trim()});
      sessionStorage.setItem('stock-sync-mapping-confirmation',JSON.stringify({id:r.id,scope:savedScope}));
      await watchConfirmation(r.id,savedScope);
    }catch(e){tell(e.message,true);}finally{busy=false;$('scan').disabled=!bridge.allowed();render();}
  }
  $('scan').onclick=scan;$('confirm').onclick=confirm;$('reviewed').onchange=counts;
  $('search').oninput=()=>{page=1;if(data)render();};$('filter').onchange=()=>{page=1;if(data)render();};
  $('prev').onclick=()=>{page--;render();};$('next').onclick=()=>{page++;render();};
  $('select').onclick=()=>{if(busy||!data)return;for(const item of data.items)if(editable(item)&&item.certain)choices.get(item.id).checked=true;changed();render();};
  $('clear').onclick=()=>{if(busy)return;for(const choice of choices.values())choice.checked=false;changed();render();};
  async function ready(){
    if(initialized)return;initialized=true;
    $('scan').disabled=!bridge.allowed();
    if(!bridge.allowed()){tell('补齐映射需要库存管理权限，请由有权限的人员处理。');return;}
    // Resume the exact submitted task after a reload; never resubmit a write.
    const pending=sessionStorage.getItem('stock-sync-mapping-confirmation');
    if(pending){try{const p=JSON.parse(pending);busy=true;$('scan').disabled=true;await watchConfirmation(p.id,p.scope);}catch(e){tell(e.message,true);}finally{busy=false;$('scan').disabled=false;}}
  }
  document.addEventListener('stock-sync-ready',ready);
  if(bridge.allowed())ready();
})();
