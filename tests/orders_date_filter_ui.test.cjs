const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.resolve(__dirname, '..');

function extractDomReadyScript(relPath, marker) {
    const html = fs.readFileSync(path.join(ROOT, relPath), 'utf8');
    const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]);
    const matched = scripts.find((code) => code.includes(marker));
    assert.ok(matched, `未找到包含 ${marker} 的脚本: ${relPath}`);
    return matched;
}

function createHarness({ formId, extraGlobals = {} }) {
    let domReadyHandler = null;
    const listeners = (target) => {
        target._listeners = target._listeners || {};
        return target._listeners;
    };
    const makeElement = (initialValue = '') => ({
        value: initialValue,
        classList: {
            _set: new Set(['btn', 'btn-quick', 'active']),
            remove(cls) {
                this._set.delete(cls);
            },
            contains(cls) {
                return this._set.has(cls);
            },
        },
        addEventListener(type, fn) {
            const map = listeners(this);
            map[type] = map[type] || [];
            map[type].push(fn);
        },
        trigger(type) {
            const map = listeners(this);
            for (const fn of map[type] || []) {
                fn();
            }
        },
    });

    const quickDateInput = makeElement('this_month');
    const dateFrom = makeElement('2026-11-01');
    const dateTo = makeElement('2026-11-16');
    const filterForm = makeElement('');
    const activeBtn = makeElement('');
    const pickerInstances = {};

    function flatpickr(inputEl, config) {
        const altInput = makeElement(inputEl.value ? `${inputEl.value}显示` : '');
        const instance = {
            input: inputEl,
            altInput,
            config,
            selectDate(dateStr) {
                inputEl.value = dateStr;
                altInput.value = dateStr;
                if (typeof config.onChange === 'function') {
                    config.onChange([new Date(dateStr)], dateStr, instance);
                }
            },
        };
        pickerInstances[inputEl === dateFrom ? 'from' : 'to'] = instance;
        return instance;
    }

    const document = {
        addEventListener(type, fn) {
            if (type === 'DOMContentLoaded') {
                domReadyHandler = fn;
            }
        },
        getElementById(id) {
            if (id === formId) return filterForm;
            if (id === 'quickDateInput') return quickDateInput;
            return null;
        },
        querySelector(sel) {
            if (sel === '#dateFrom') return dateFrom;
            if (sel === '#dateTo') return dateTo;
            return null;
        },
        querySelectorAll(sel) {
            if (sel === '.quick-filters .btn-quick.active') {
                return activeBtn.classList.contains('active') ? [activeBtn] : [];
            }
            return [];
        },
    };

    const sandbox = {
        document,
        window: { flatpickr, localStorage: { getItem: () => null, setItem: () => {} } },
        flatpickr,
        localStorage: { getItem: () => null, setItem: () => {} },
        ...extraGlobals,
    };
    return {
        sandbox,
        runReady() {
            assert.ok(domReadyHandler, 'DOMContentLoaded 未注册');
            domReadyHandler();
        },
        quickDateInput,
        dateFrom,
        dateTo,
        filterForm,
        activeBtn,
        pickerInstances,
    };
}

test('orders.html 选择 10 月自定义日期后自动清空 quick_date 隐藏域与按钮高亮', () => {
    const script = extractDomReadyScript('templates/orders.html', 'ordersFilterForm');
    const h = createHarness({ formId: 'ordersFilterForm' });
    vm.runInNewContext(script, h.sandbox);
    h.runReady();

    assert.equal(h.quickDateInput.value, 'this_month');
    assert.equal(h.activeBtn.classList.contains('active'), true);

    h.pickerInstances.from.selectDate('2026-10-01');
    h.pickerInstances.to.selectDate('2026-10-31');

    assert.equal(h.dateFrom.value, '2026-10-01');
    assert.equal(h.dateTo.value, '2026-10-31');
    assert.equal(h.quickDateInput.value, '');
    assert.equal(h.activeBtn.classList.contains('active'), false);
});

test('orders.html 清空可见日期输入框并失焦或提交时同步清空底层日期值', () => {
    const script = extractDomReadyScript('templates/orders.html', 'ordersFilterForm');
    const h = createHarness({ formId: 'ordersFilterForm' });
    vm.runInNewContext(script, h.sandbox);
    h.runReady();

    h.pickerInstances.from.altInput.value = '   ';
    h.pickerInstances.from.altInput.trigger('input');
    h.pickerInstances.from.altInput.trigger('blur');

    h.pickerInstances.to.altInput.value = '';
    h.filterForm.trigger('submit');

    assert.equal(h.dateFrom.value, '');
    assert.equal(h.dateTo.value, '');
    assert.equal(h.quickDateInput.value, '');
});

test('orders.html 未改动日期直接提交筛选时保留当前 quick_date', () => {
    const script = extractDomReadyScript('templates/orders.html', 'ordersFilterForm');
    const h = createHarness({ formId: 'ordersFilterForm' });
    vm.runInNewContext(script, h.sandbox);
    h.runReady();

    h.filterForm.trigger('submit');
    assert.equal(h.quickDateInput.value, 'this_month');
    assert.equal(h.dateFrom.value, '2026-11-01');
    assert.equal(h.dateTo.value, '2026-11-16');
    assert.equal(h.activeBtn.classList.contains('active'), true);
});

test('customers.html 选择自定义日期时同样清空 quick_date 隐藏域与按钮高亮', () => {
    const html = fs.readFileSync(path.join(ROOT, 'templates/customers.html'), 'utf8');
    const match = html.match(/\/\/ 日期选择器初始化及快捷预设联动清理[\s\S]*?if \(filterForm\) \{[\s\S]*?\}\s*\}\s*\n/);
    assert.ok(match, '未找到 customers.html 日期选择器绑定代码块');
    const wrapped = `${match[0]}\n});`;
    const h = createHarness({ formId: 'customersFilterForm' });
    vm.runInNewContext(wrapped, h.sandbox);
    h.runReady();

    h.pickerInstances.from.selectDate('2026-10-01');
    h.pickerInstances.to.selectDate('2026-10-31');

    assert.equal(h.quickDateInput.value, '');
    assert.equal(h.activeBtn.classList.contains('active'), false);
});

