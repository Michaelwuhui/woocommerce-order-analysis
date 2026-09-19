/* Admin stocktake entry from a fulfillment; all authority is checked server-side. */
function fulfillmentStocktakePayload(context, quantities, note, requestKey) {
  if (!String(note || '').trim()) throw new Error('请填写库存调整原因');
  if (String(note).trim().length > 2000) throw new Error('调整原因不能超过 2000 字');
  const items = [];
  for (const item of context.items) {
    const raw = String(quantities[item.sku_id] ?? '').trim();
    const qty = Number(raw);
    if (!/^\d+$/.test(raw) || !Number.isSafeInteger(qty) || qty > 100000000) {
      throw new Error(`${item.sku_code}：实盘总数必须是有效的非负整数`);
    }
    if (qty < item.reserved) throw new Error(`${item.sku_code}：实盘不能低于已预留的 ${item.reserved} 件`);
    if (qty !== item.on_hand) items.push({sku_id: item.sku_id, qty,
      baseline: item.on_hand, baseline_reserved: item.reserved, baseline_movement: item.movement});
  }
  if (!items.length) throw new Error('请先修改需要调整的商品数量');
  return {request_key: requestKey, revision: context.revision, note: String(note).trim(), items};
}

function fulfillmentStocktakeButtons(targets, orderId, permitted, escape) {
  if (!permitted) return '';
  return (targets || []).map(target => `<button type="button" class="btn btn-outline-warning btn-sm" data-ff-stock-id="${escape(target.fulfillment_id)}" data-ff-stock-order="${escape(orderId)}">快速调整库存${targets.length > 1 ? ' · ' + escape(target.warehouse_name) : ''}</button>`).join('');
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {fulfillmentStocktakePayload, fulfillmentStocktakeButtons};
}

if (typeof document !== 'undefined') {
  let draft = null;
  let loadVersion = 0;
  const panel = () => document.getElementById('ffStockAdjustmentPanel');
  const endpoint = id => '/api/inv/operations/fulfillment/' + encodeURIComponent(id) + '/stocktake';
  const message = (text, danger = true) => {
    const node = document.getElementById('ffStockMessage');
    if (node) { node.className = 'alert ' + (danger ? 'alert-danger' : 'alert-success'); node.textContent = text; }
  };

  async function openAdjustment(id, orderId) {
    const host = panel();
    if (!host) return;
    const version = ++loadVersion;
    draft = null;
    host.innerHTML = '<div class="alert alert-secondary">正在读取本仓库存…</div>';
    host.scrollIntoView({block: 'nearest'});
    try {
      const context = await ffJson(endpoint(id));
      if (version !== loadVersion || !host.isConnected) return;
      draft = {id, orderId, context, payload: null, busy: false};
      const missing = context.unavailable.length
        ? `<div class="alert alert-warning">以下商品未映射或未纳入本仓库存，不能在此调整：${context.unavailable.map(i => esc(i.name)).join('、')}</div>` : '';
      host.innerHTML = `<section class="card border-warning bg-dark mb-3"><div class="card-header d-flex flex-wrap justify-content-between gap-2"><b>快速调整库存 · ${esc(context.warehouse_name)} · #${esc(context.order_number)}</b><button type="button" class="btn btn-sm btn-outline-light" id="ffStockClose">关闭调整</button></div>
        <form id="ffStockForm" class="card-body"><p class="text-white-50">填写实盘总数，包含已预留但尚未出库的商品。保存立即入账，预留数量保持不变；实盘不能低于预留。</p>${missing}
        <div class="table-responsive"><table class="table table-dark table-sm align-middle"><thead><tr><th>商品 / SKU</th><th>本单数量 / 缺货</th><th>现存 / 预留 / 可用</th><th>实盘总数</th><th>增减</th></tr></thead><tbody>${context.items.map(item => `<tr><td>${esc(item.name)}<div class="small text-white-50">${esc(item.sku_code)}</div></td><td>${item.ordered_qty} / ${item.shortage_qty}</td><td>${item.on_hand} / ${item.reserved} / ${item.available}</td><td><input id="ffStockQty${item.sku_id}" data-stock-sku="${item.sku_id}" class="form-control form-control-sm" style="min-width:100px" type="number" step="1" min="${item.reserved}" max="100000000" value="${item.on_hand}" required aria-label="${esc(item.sku_code)} 实盘总数"></td><td id="ffStockDelta${item.sku_id}">0</td></tr>`).join('') || '<tr><td colspan="5">本仓没有可调整的订单商品。</td></tr>'}</tbody></table></div>
        <label for="ffStockNote" class="form-label">调整原因</label><textarea id="ffStockNote" class="form-control mb-3" rows="2" maxlength="2000" required placeholder="填写盘点依据和差异原因"></textarea><div id="ffStockMessage" role="status"></div>
        <div class="d-flex flex-wrap gap-2"><button type="submit" class="btn btn-warning" id="ffStockSave" ${context.items.length ? '' : 'disabled'}>确认调整库存</button><button type="button" class="btn btn-outline-light" id="ffStockReload">重新读取库存</button></div></form></section>`;
    } catch (error) {
      if (version === loadVersion && host.isConnected) host.innerHTML = `<div class="alert alert-danger">${esc(error.message)}</div>`;
    }
  }

  document.getElementById('ffDetail').addEventListener('click', event => {
    const button = event.target.closest('[data-ff-stock-id]');
    if (button) { openAdjustment(button.dataset.ffStockId, button.dataset.ffStockOrder); return; }
    if (draft?.busy) return;
    if (event.target.id === 'ffStockClose') { ++loadVersion; draft = null; panel().innerHTML = ''; }
    if (event.target.id === 'ffStockReload' && draft) openAdjustment(draft.id, draft.orderId);
  });
  document.getElementById('ffDetail').addEventListener('input', event => {
    const sid = Number(event.target.dataset.stockSku);
    if (!sid || !draft) return;
    const item = draft.context.items.find(item => item.sku_id === sid);
    const delta = Number(event.target.value) - item.on_hand;
    const node = document.getElementById('ffStockDelta' + sid);
    node.textContent = event.target.value === '' || !Number.isInteger(delta) ? '—' : (delta > 0 ? '+' : '') + delta;
    node.className = delta ? 'text-warning fw-bold' : '';
  });
  document.getElementById('ffDetail').addEventListener('submit', async event => {
    if (event.target.id !== 'ffStockForm') return;
    event.preventDefault();
    const current = draft;
    if (!current || current.busy) return;
    const form = event.target;
    try {
      if (!current.payload) {
        const quantities = Object.fromEntries(current.context.items.map(item =>
          [item.sku_id, document.getElementById('ffStockQty' + item.sku_id).value]));
        const payload = fulfillmentStocktakePayload(current.context, quantities,
          document.getElementById('ffStockNote').value, 'ff-stock-' + crypto.randomUUID());
        if (!confirm(`确认调整 ${current.context.warehouse_name} 的 ${payload.items.length} 项库存？将按实盘总数立即入账，并生成盘点记录。`)) return;
        current.payload = payload;
      }
      current.busy = true;
      for (const control of form.querySelectorAll('input,textarea,button')) control.disabled = true;
      document.getElementById('ffStockClose').disabled = true;
      message('正在保存库存调整…', false);
      const result = await ffJson(endpoint(current.id), {method: 'POST', body: JSON.stringify(current.payload)});
      if (result.status !== 'approved' || !Number.isSafeInteger(result.id) || result.id <= 0) {
        throw new Error('尚未收到有效的库存入账确认');
      }
      const text = `库存已调整，盘点单 #${result.id}。库存增加后，系统会按现有规则重新核对缺货分配。`;
      if (draft === current && form.isConnected) {
        draft = null;
        await showFulfillment(current.orderId);
        if (panel()) panel().innerHTML = `<div class="alert alert-success">${esc(text)} <a class="alert-link" href="/inventory/operations?kind=stocktake">查看盘点记录</a></div>`;
      }
      ffAlert(text, false);
      await loadFulfillments();
    } catch (error) {
      if (draft !== current || !form.isConnected) { ffAlert(error.message); return; }
      const uncertain = current.payload && (!error.status || error.status >= 500);
      if (!uncertain) current.payload = null;
      for (const control of form.querySelectorAll('input,textarea,button')) control.disabled = !!uncertain;
      const save = document.getElementById('ffStockSave');
      save.disabled = false;
      save.textContent = uncertain ? '重试核对本次调整' : '确认调整库存';
      document.getElementById('ffStockClose').disabled = !!uncertain;
      message(error.message + (uncertain ? '；结果尚未确认，请重试核对本次调整，不要重复新建。' : ''));
    } finally { current.busy = false; }
  });
}
