const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..');

function loadModScriptContext() {
    const html = fs.readFileSync(path.join(ROOT, 'templates/base.html'), 'utf8');
    const startIdx = html.indexOf('// ---- 订单详情：更换口味 / 修改商品逻辑 ----');
    assert.ok(startIdx > 0, '未找到 base.html 中的商品修改前端代码块');
    const endMarker = '// Open the per-email detail modal.';
    const endIdx = html.indexOf(endMarker, startIdx);
    assert.ok(endIdx > startIdx, '未找到商品修改代码块结束位置');
    const code = html.slice(startIdx, endIdx);

    const elements = {};
    const makeEl = (id) => ({
        id,
        value: '',
        textContent: '',
        innerHTML: '',
        checked: false,
        disabled: false,
        min: '1',
        max: '1',
        classList: {
            _s: new Set(['d-none']),
            add(c) { this._s.add(c); },
            remove(c) { this._s.delete(c); },
            toggle(c, force) {
                if (force === undefined) {
                    if (this._s.has(c)) this._s.delete(c); else this._s.add(c);
                } else if (force) {
                    this._s.add(c);
                } else {
                    this._s.delete(c);
                }
            },
            contains(c) { return this._s.has(c); },
        },
    });

    const ids = [
        'modItemOrderBadge', 'modOrigItemName', 'modOrigItemSku', 'modOrigItemQty', 'modOrigItemTotal',
        'modReplaceQtyInput', 'modQtyMaxHint', 'modFlavorFilterInput', 'modProductSearchInput',
        'modPriceKeepRadio', 'modPriceSyncRadio', 'modErrorAlertBox', 'modifyOrderItemModal',
        'modModeSameBtn', 'modModeOtherBtn', 'modProductSearchBox', 'modSameProductPriceNotice',
        'modCrossProductPriceRadios', 'modSelectedProductTitle', 'modFlavorOptionsList',
        'modProductSearchResults', 'confirmModifyItemBtn', 'modPreviewSummaryBox', 'modSplitHintText',
    ];
    for (const id of ids) {
        elements[id] = makeEl(id);
    }

    const fetchCalls = [];
    let detailReloadedOrderId = null;

    async function fakeFetch(url, options) {
        fetchCalls.push({ url, options });
        if (url.includes('/item-options?product_id=100')) {
            return {
                json: async () => ({
                    success: true,
                    currency: 'PLN',
                    product: { product_id: 100, name: 'Fumot 15000', type: 'variable' },
                    options: [
                        {
                            option_key: '100:1001:',
                            product_id: 100,
                            variation_id: 1001,
                            flavor_label: 'MIXED BERRY',
                            full_name: 'Fumot 15000 - MIXED BERRY',
                            sku: 'F-MIXED',
                            price: '64.90',
                            in_stock: true,
                            stock_quantity: 10,
                            explicit_meta: [],
                        },
                        {
                            option_key: '100:1002:',
                            product_id: 100,
                            variation_id: 1002,
                            flavor_label: 'WATERMELON ICE',
                            full_name: 'Fumot 15000 - WATERMELON ICE',
                            sku: 'F-WATER',
                            price: '64.90',
                            in_stock: true,
                            stock_quantity: 5,
                            explicit_meta: [],
                        },
                    ],
                }),
            };
        }
        if (url.includes('/item-options?product_id=200')) {
            return {
                json: async () => ({
                    success: true,
                    currency: 'PLN',
                    product: { product_id: 200, name: 'Blade 30000', type: 'variable' },
                    options: [
                        {
                            option_key: '200:2001:',
                            product_id: 200,
                            variation_id: 2001,
                            flavor_label: 'BLUE RAZZ',
                            full_name: 'Blade 30000 - BLUE RAZZ',
                            sku: 'B-BLUE',
                            price: '89.90',
                            in_stock: true,
                            stock_quantity: 8,
                            explicit_meta: [],
                        },
                    ],
                }),
            };
        }
        if (url.includes('/search-products')) {
            return {
                json: async () => ({
                    success: true,
                    currency: 'PLN',
                    products: [{ product_id: 200, name: 'Blade 30000', type: 'variable', price: '89.90', in_stock: true, variation_count: 1 }],
                }),
            };
        }
        if (url.includes('/modify-item')) {
            return {
                status: 200,
                text: async () => JSON.stringify({ success: true, message: '修改成功' }),
            };
        }
        throw new Error('Unexpected URL: ' + url);
    }

    const windowObj = {
        _currentDetailOrder: {
            id: '1-9002',
            number: '9002',
            currency: 'PLN',
            total: '120.00',
            can_modify_items: true,
            line_items: [
                {
                    id: 801,
                    product_id: 100,
                    variation_id: 1001,
                    name: 'Fumot 15000 - MIXED BERRY',
                    display_name: 'Fumot 15000 - MIXED BERRY',
                    sku: 'F-MIXED',
                    quantity: 3,
                    price: 33.33,
                    total: '100.00',
                },
            ],
        },
        _orderItemModifiedReloadNeeded: false,
    };

    const sandbox = {
        window: windowObj,
        document: {
            getElementById(id) { return elements[id] || null; },
        },
        bootstrap: {
            Modal: {
                getOrCreateInstance: () => ({ show() {} }),
                getInstance: () => ({ hide() {} }),
            },
        },
        fetch: fakeFetch,
        encodeURIComponent,
        parseInt,
        parseFloat,
        isNaN,
        Math,
        Number,
        String,
        JSON,
        showToast: () => {},
        alert: () => {},
        showOrderDetail: (oid) => { detailReloadedOrderId = oid; },
    };

    vm.runInNewContext(code, sandbox);
    return { sandbox, elements, fetchCalls, getReloadedId: () => detailReloadedOrderId };
}

test('同款拆分更换口味：自动标记当前原口味、实时计算零尾差拆分预览', async () => {
    const { sandbox, elements } = loadModScriptContext();
    sandbox.openModifyOrderItemModal(0);
    await new Promise((r) => setImmediate(r));

    assert.ok(elements.modFlavorOptionsList.innerHTML.includes('当前原口味'));
    assert.ok(elements.modFlavorOptionsList.innerHTML.includes('WATERMELON ICE'));

    sandbox.selectModOptionByKey('100:1002:');
    elements.modReplaceQtyInput.value = '1';
    sandbox.updateModPricePreview();

    assert.equal(elements.confirmModifyItemBtn.disabled, false);
    assert.ok(elements.modSplitHintText.textContent.includes('原口味保留 2 件 + 新口味 1 件'));
    assert.ok(elements.modPreviewSummaryBox.innerHTML.includes('66.67 PLN'));
    assert.ok(elements.modPreviewSummaryBox.innerHTML.includes('33.33 PLN'));
    assert.ok(elements.modPreviewSummaryBox.innerHTML.includes('订单总金额保持 120.00 PLN 不变'));
});

test('跨款更换产品：支持切换同步新产品售价并提交完整修改请求', async () => {
    const { sandbox, elements, fetchCalls, getReloadedId } = loadModScriptContext();
    sandbox.openModifyOrderItemModal(0);
    await new Promise((r) => setImmediate(r));

    sandbox.switchModifyItemMode('other');
    await new Promise((r) => setImmediate(r));
    assert.equal(elements.modCrossProductPriceRadios.classList.contains('d-none'), false);

    sandbox.selectSearchedProductForMod(200);
    await new Promise((r) => setImmediate(r));
    sandbox.selectModOptionByKey('200:2001:');

    // 拆分 1 件并选择同步新产品售价（保留 2 件 66.67 + 新产品 1 件 89.90 = 156.57，比原 100.00 增加 +56.57，订单总计 176.57）
    elements.modReplaceQtyInput.value = '1';
    elements.modPriceKeepRadio.checked = false;
    elements.modPriceSyncRadio.checked = true;
    sandbox.updateModPricePreview();

    assert.ok(elements.modPreviewSummaryBox.innerHTML.includes('+56.57 PLN'));
    assert.ok(elements.modPreviewSummaryBox.innerHTML.includes('176.57 PLN'));

    sandbox.submitModifyOrderItem();
    await new Promise((r) => setImmediate(r));

    const postCall = fetchCalls.find((c) => c.url.endsWith('/modify-item'));
    assert.ok(postCall, '未发起 /modify-item 请求');
    const body = JSON.parse(postCall.options.body);
    assert.deepEqual(body, {
        line_item_id: 801,
        expected_original_qty: 3,
        replace_qty: 1,
        target_product_id: 200,
        target_variation_id: 2001,
        target_flavor: 'BLUE RAZZ',
        price_mode: 'sync',
    });
    assert.equal(getReloadedId(), '1-9002');
    assert.equal(sandbox.window._orderItemModifiedReloadNeeded, true);
});
