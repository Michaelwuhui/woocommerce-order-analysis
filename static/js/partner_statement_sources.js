/* Frozen product-source bill: summary -> products -> original order/parcel. */
(function (root) {
    'use strict';
    const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
    const amount = value => value == null ? '待核实' : Number(value).toLocaleString('zh-CN', {minimumFractionDigits:2,maximumFractionDigits:2});
    const card = (title, value, unit, hint, color) => `<div class="col-md-4"><div class="card h-100 border-secondary" style="background:${color || '#17273e'}"><div class="card-body"><div class="text-white-50 small">${esc(title)}</div><div class="h4 text-white my-2">${amount(value)} <small class="fs-6">${esc(unit)}</small></div><div class="small text-white-50">${esc(hint)}</div></div></div></div>`;
    const labels = {partner_stock:'合伙人自备货',transit_service:'我方备货 · 中转服务',pending:'来源待核实'};
    function summary(snapshot, live) {
        if (!snapshot?.settlement) return '';
        const s = snapshot.settlement, curr = snapshot.currency || 'PLN';
        const groupAmount = (row,key) => row.mode==='pending' && row.pending_count ? '<span class="text-warning">待核实</span>' : amount(row[key]);
        return `<section class="source-summary mb-3" data-source-summary>
            <div class="d-flex align-items-center flex-wrap gap-2 mb-3"><h5 class="mb-0 text-white">按商品来源核算</h5><span class="badge bg-info text-dark">${live ? '实时参考' : `第 ${esc(snapshot.revision)} 版 · 冻结快照`}</span><span class="badge ${s.ready ? 'bg-success' : 'bg-warning text-dark'}">${esc(s.amounts_label)}</span></div>
            <div class="row g-2 mb-3">
                ${card('自备货商品利润 · 合伙人',s.partner_profit,curr,'只计算合伙人自己备货的商品','#23332b')}
                ${card('中转打包费 · 应付合伙人',s.packing_fee_cny,'CNY',`${s.packing_basis==='shipped_unit' ? s.packing_units+' 支'+(s.packing_unit_fee_cny==null?'':' × '+s.packing_unit_fee_cny+' 元/支')+' · '+s.packing_charges+' 个包裹' : s.packing_charges+' 次收费'} · 按实际出库月 · 不自动抵扣`,'#17324a')}
                ${card('商品货款应收 · 我方',s.our_receivable,curr,'自备货应收 + 中转商品应交回；实收回款另核','#342b1b')}
            </div>
            <div class="table-responsive"><table class="table table-dark table-sm align-middle mb-2"><thead><tr><th>商品供货方式</th><th class="text-end">涉及订单 / 销售件数</th><th class="text-end">商品销售</th><th class="text-end">返还合伙人成本</th><th class="text-end">合伙人商品利润</th><th class="text-end">商品应收</th></tr></thead><tbody>
                ${(snapshot.groups || []).map(g => `<tr><td><span class="badge ${g.mode === 'transit_service' ? 'bg-info text-dark' : g.mode === 'pending' ? 'bg-warning text-dark' : 'bg-success'}">${esc(g.label)}</span>${g.mode === 'transit_service' ? '<div class="small text-info mt-1">不返还备货成本 · 不参与商品分成</div>' : ''}</td><td class="text-end">${g.orders} 单 / ${g.quantity} 件</td><td class="text-end">${amount(g.revenue)}</td><td class="text-end">${groupAmount(g,'cost_recovery')}</td><td class="text-end">${groupAmount(g,'partner_profit')}</td><td class="text-end">${groupAmount(g,'our_receivable')}</td></tr>`).join('')}
            </tbody></table></div>
            <div class="small text-white-50">商品金额单位：${esc(curr)}。混合订单在各供货方式分别出现，订单数不直接相加。商品结算按下单月，打包费按实际出库月；本账只核对本合伙人绑定站点和币种。</div>
            <div class="small text-white-50 mt-2">中转退件运费损失 ${amount(s.transit_shipping_loss_separate)} ${esc(curr)}，由我方另账核算；打包费与商品应收分别列示。${esc(s.packing_policy || s.reship_policy || '普通补发收费，仓库漏发免收')}。</div>
            ${!s.ready ? `<div class="alert alert-warning py-2 mt-2 mb-0">${s.pending_orders} 单商品来源或成本待核实，${s.packing_pending} 条服务记录待核实。当前为已核实行小计，待补全后重新生成才能确认或锁定。</div>` : ''}
        </section>`;
    }
    function render(statement) {
        return summary(statement.source_snapshot, false)+`
            <section class="card data-card mb-3" id="source-detail-panel">
                <div class="card-header d-flex flex-wrap gap-2 justify-content-between align-items-center"><strong class="text-white">汇总与原始明细追溯</strong><span class="small text-white-50">同一冻结版本 · 分页读取</span></div>
                <div class="card-body">
                    <div class="d-flex flex-wrap gap-2 mb-3"><button class="btn btn-sm btn-info" data-source-kind="products">按产品汇总</button><button class="btn btn-sm btn-outline-info" data-source-kind="lines">订单商品来源</button><button class="btn btn-sm btn-outline-info" data-source-kind="fees">中转打包费流水</button><button class="btn btn-sm btn-outline-light" data-source-kind="versions">修改前记录</button></div>
                    <div class="d-flex flex-wrap gap-2 mb-3" data-source-filters><select class="form-select form-select-sm bg-dark text-white border-secondary" style="width:auto" data-source-mode><option value="">全部供货方式</option><option value="partner_stock">合伙人自备货</option><option value="transit_service">我方备货 · 中转服务</option><option value="pending">来源待核实</option></select><input class="form-control form-control-sm bg-dark text-white border-secondary" style="flex:1;min-width:180px" data-source-search placeholder="订单号、商品、SKU、站点或运单号"><button class="btn btn-sm btn-outline-info" data-source-search-button>查询</button><label class="form-check-label small text-white-50 align-self-center"><input type="checkbox" data-source-pending> 只看待核实</label></div>
                    <div data-source-content aria-live="polite">正在读取冻结明细…</div>
                    <div class="d-flex justify-content-between align-items-center mt-3"><span class="small text-white-50" data-source-count></span><div class="d-flex gap-2"><button class="btn btn-sm btn-outline-light" data-source-prev>上一页</button><button class="btn btn-sm btn-outline-light" data-source-next>下一页</button></div></div>
                    <div class="mt-3" data-source-decision-save hidden><button class="btn btn-sm btn-primary" data-source-save>保存补发分类并生成新版账单</button><span class="small text-white-50 ms-2">保留原版；需要填写核对依据。</span></div>
                </div>
            </section>`;
    }
    function table(headers, contents) {
        return `<div class="table-responsive"><table class="table table-dark table-sm align-middle"><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${contents.length ? contents.map(c=>`<tr>${c.map(v=>`<td>${v}</td>`).join('')}</tr>`).join('') : `<tr><td colspan="${headers.length}" class="text-center text-white-50 py-4">没有符合条件的记录</td></tr>`}</tbody></table></div>`;
    }
    function renderRows(data, kind, editable) {
        const values = data.rows || [];
        if (kind === 'products') return table(['供货方式','商品 / 仓库','订单 / 件数','商品销售','返还合伙人成本','合伙人利润','商品应收','核算状态'],values.map(r=>[
            esc(r.mode_label),`${esc(r.product_name)}<div class="small text-white-50">${esc(r.warehouse_name)}</div>`,`${r.orders} 单 / ${r.quantity} 件`,amount(r.revenue),amount(r.cost_recovery),amount(r.partner_profit),amount(r.our_receivable),r.pending_count ? '<span class="text-warning">待核实</span>' : '已核算']));
        if (kind === 'lines') return table(['订单 / 站点','商品','供货方式 / 仓库','来源归属 / OMS 实发数量','商品销售','返还合伙人成本','合伙人利润','商品应收','来源与成本凭证'],values.map(r=>[
            `<a href="/orders?search=${encodeURIComponent(r.order_number || r.order_id)}" target="_blank" rel="noopener">${esc(r.order_number || r.order_id)}</a><div class="small text-white-50">${esc(r.source)}<br>${esc(r.order_date)}<br>${esc(r.order_status)}</div>`,
            `${esc(r.product_name)}<div class="small text-white-50">商品行 ${esc(r.line_item_id)}${r.sku ? ' · '+esc(r.sku) : ''}</div>`,
            `${esc(r.mode_label)}<div class="small text-white-50">${esc(r.warehouse_name)}</div>`,`${r.quantity} / ${r.fulfilled_quantity || 0}`,
            amount(r.revenue),amount(r.cost_recovery),amount(r.partner_profit),amount(r.our_receivable),
            `<div class="small">${esc(r.source_basis)}${r.fulfillment_id ? '<br>履约单 '+esc(r.fulfillment_id) : ''}${r.source_override_id ? '<br>来源确认 '+esc(r.source_override_id)+'<br>'+esc(r.ownership_evidence) : ''}${r.cost_evidence?.cost_id ? '<br>成本 #'+esc(r.cost_evidence.cost_id)+' · '+esc(r.cost_evidence.effective_date) : ''}${r.pending?.length ? '<br><span class="text-warning">'+esc(r.pending.join('；'))+'</span>' : ''}</div>`]));
        return table(['订单 / 站点','运单 / 实际出库','包裹中转商品 / 实发支数','计费依据','打包费 CNY','费用规则 / 核对依据'],values.map(r=>{
            const perUnit = r.billing_basis === 'shipped_unit';
            const canDecide = !perUnit && editable && (r.kind === 'reship' || (r.pending || []).some(p=>p.includes('原发/补发')));
            const decision = r.decision || {};
            const control = canDecide ? `<div class="mt-2"><select class="form-select form-select-sm bg-dark text-white border-secondary" data-reship-decision="${esc(r.shipment_id)}"><option value="unknown">选择核对结果</option>${[['original','原订单首次发货'],['ordinary','普通补发 · 收费'],['warehouse_omission','仓库漏发补发 · 免收']].map(([key,label])=>`<option value="${key}" ${decision.reason===key?'selected':''}>${label}</option>`).join('')}</select><input class="form-control form-control-sm mt-1 bg-dark text-white border-secondary" data-reship-evidence="${esc(r.shipment_id)}" placeholder="填写核对依据" value="${esc(decision.evidence || '')}"></div>` : '';
            return [`${esc(r.order_number || r.order_id)}<div class="small text-white-50">${esc(r.source)}<br>${esc(r.order_status || '')}</div>`,`${esc(r.tracking_number)}<div class="small text-white-50">${esc(r.shipped_at)}<br>${esc(r.shipment_id)}</div>`,
                perUnit ? (r.items || []).map(i=>`<div class="small">${esc(i.product_name)} · ${esc(i.quantity)} 支<br>商品行 ${esc(i.line_item_id)} / 履约商品 ${esc(i.fulfillment_item_id)}</div>`).join('') || '实发商品待核实' : '历史按订单计费',
                `${perUnit ? (r.quantity==null?'支数待核实':esc(r.quantity)+' 支')+' × '+amount(r.unit_fee_cny)+' CNY/支' : r.kind==='reship'?'补发':'原订单'}<div class="small text-white-50">${esc(r.reason)}</div>`,amount(r.amount_cny),`<div class="small text-white-50">${esc(r.rule_name)}<br>${esc(r.rule_id)}${r.pending?.length ? '<br><span class="text-warning">'+esc(r.pending.join('；'))+'</span>' : ''}</div>${control}`];
        }));
    }
    function mount(body, statement, options) {
        const panel = body.querySelector('#source-detail-panel');
        if (!panel) return;
        let kind = 'products', page = 1, serial = 0, decisions = {};
        const $ = selector => panel.querySelector(selector);
        const editable = options.editable && statement.status === 'generated';
        function remember() {
            panel.querySelectorAll('[data-reship-decision]').forEach(el=>{
                const evidence = panel.querySelector(`[data-reship-evidence="${CSS.escape(el.dataset.reshipDecision)}"]`);
                if (el.value !== 'unknown') decisions[el.dataset.reshipDecision] = {reason:el.value,evidence:evidence.value.trim()};
            });
        }
        async function load() {
            remember();
            const request = ++serial;
            $('[data-source-content]').innerHTML = '<div class="text-white-50 py-3">读取冻结明细…</div>';
            $('[data-source-prev]').disabled = true; $('[data-source-next]').disabled = true;
            $('[data-source-filters]').hidden = kind === 'versions';
            $('[data-source-decision-save]').hidden = !(editable && kind === 'fees' && statement.source_snapshot.settlement.packing_basis!=='shipped_unit');
            try {
                if (kind === 'versions') {
                    const data = await options.api(`/api/reconciliation/statements/${statement.id}/versions`);
                    if (request !== serial || !panel.isConnected) return;
                    $('[data-source-content]').innerHTML = table(['版本 / 时间','本版商品应收','更正前净额 / 成本','更正前合伙人利润 / 我方应收','保留原订单'],(data.versions || []).map(v=>[
                        `第 ${v.revision} 版<div class="small text-white-50">${esc(v.created_at)}</div>`,amount(v.settlement?.our_receivable),
                        v.previous_statement ? `${amount(v.previous_statement.total_net_pln)} / ${amount(v.previous_statement.cost_amount_pln)}` : '首次生成',
                        v.previous_statement ? `${amount(v.previous_statement.partner_profit_pln)} / ${amount(v.previous_statement.our_receivable_pln)}` : '—',String(v.previous_order_count)]));
                    $('[data-source-count]').textContent = '历史金额与原订单快照已保留；更正不修改已确认账单。';
                    return;
                }
                const params = new URLSearchParams({kind,page,per_page:50,revision:statement.source_snapshot.revision});
                if ($('[data-source-mode]').value && kind !== 'fees') params.set('mode',$('[data-source-mode]').value);
                if ($('[data-source-pending]').checked) params.set('pending','1');
                if ($('[data-source-search]').value.trim()) params.set('search',$('[data-source-search]').value.trim());
                const data = await options.api(`/api/reconciliation/statements/${statement.id}/source-detail?${params}`);
                if (request !== serial || !panel.isConnected) return;
                for (const row of data.rows || []) if (decisions[row.shipment_id]) row.decision = decisions[row.shipment_id];
                $('[data-source-content]').innerHTML = renderRows(data,kind,editable);
                $('[data-source-count]').textContent = `第 ${page} / ${Math.max(1,Math.ceil(data.count/50))} 页 · 共 ${data.count} 条 · 第 ${data.revision} 版`;
                $('[data-source-prev]').disabled = page <= 1; $('[data-source-next]').disabled = page*50 >= data.count;
            } catch (error) {
                if (request===serial && panel.isConnected) $('[data-source-content]').innerHTML = `<div class="text-danger">${esc(error.message)}</div>`;
            }
        }
        panel.querySelectorAll('[data-source-kind]').forEach(button=>button.addEventListener('click',()=>{kind=button.dataset.sourceKind;page=1;load();}));
        $('[data-source-mode]').addEventListener('change',()=>{page=1;load();});
        $('[data-source-pending]').addEventListener('change',()=>{page=1;load();});
        $('[data-source-search-button]').addEventListener('click',()=>{page=1;load();});
        $('[data-source-search]').addEventListener('keydown',event=>{if(event.key==='Enter'){page=1;load();}});
        $('[data-source-prev]').addEventListener('click',()=>{page--;load();});
        $('[data-source-next]').addEventListener('click',()=>{page++;load();});
        $('[data-source-save]').addEventListener('click',async()=>{
            remember();
            if (!Object.keys(decisions).length) return options.toast('请先选择需要核对的发货或补发分类','warning');
            try {
                $('[data-source-save]').disabled = true;
                await options.api('/api/reconciliation/statements/generate',{method:'POST',body:JSON.stringify({partner_id:statement.partner_id,year:statement.period_year,month:statement.period_month,mode:statement.calc_mode,reship_decisions:decisions})});
                options.toast('分类已保存，新版账单已生成，原版已保留','success');
                options.refresh();
            } catch(error) { options.toast(error.message,'danger'); }
            finally { if(panel.isConnected) $('[data-source-save]').disabled = false; }
        });
        load();
    }
    const api = {esc,amount,summary,render,renderRows,mount};
    root.PartnerStatementSources = api;
    if (typeof module !== 'undefined') module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
