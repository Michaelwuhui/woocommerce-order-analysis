/* Order-level shortage lines can be absent from every warehouse allocation. */
function fulfillmentShortageSummary(items) {
  return (items || []).map(item => `${item.sku_code || item.name} 缺 ${item.shortage_qty} 件`).join('；');
}

function fulfillmentShortageDetails(items, escape) {
  if (!items?.length) return '';
  const rows = items.map(item => `<tr><td>${escape(item.name)}</td><td>${escape(item.sku_code || '未映射')}</td><td>${item.ordered_qty}</td><td>${item.allocated_qty}</td><td class="text-danger fw-bold">${item.shortage_qty}</td><td>${item.sku_id ? '库存不足 / 待重新分配' : '未建立 SKU 映射'}</td></tr>`).join('');
  return `<section class="card border-danger bg-dark mb-3" id="ffShortageDetails"><div class="card-header text-danger fw-bold">订单缺货明细</div><div class="card-body"><p>以下商品尚未全部分配到仓库，完全未分配的商品也包含在内。</p><div class="table-responsive"><table class="table table-dark table-sm"><thead><tr><th>商品</th><th>SKU</th><th>本单数量</th><th>已分配</th><th>仍缺</th><th>原因</th></tr></thead><tbody>${rows}</tbody></table></div><div class="small text-white-50">库存入账后，缺货数量以后台重新分配的结果为准。</div></div></section>`;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {fulfillmentShortageSummary, fulfillmentShortageDetails};
}
