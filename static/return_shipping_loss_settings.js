(() => {
    const root = document.getElementById('returnShippingLossSettings');
    if (!root) return;
    const body = document.getElementById('returnLossRulesBody');
    const status = document.getElementById('returnLossSettingsStatus');
    const add = document.getElementById('returnLossAddBtn');
    const save = document.getElementById('returnLossSaveBtn');
    const reload = document.getElementById('returnLossReloadBtn');
    let version = null, warehouses = [], busy = false;
    const say = (message, error = false) => {
        status.textContent = message;
        status.className = 'small mt-2 ' + (error ? 'text-danger' : 'text-info');
    };
    function lock(value) {
        busy = value;
        root.querySelectorAll('input,select,button').forEach(el => { el.disabled = value; });
        add.disabled = save.disabled = value || version === null;
    }
    function addRow(rule = {}) {
        const row = document.createElement('tr');
        function cell() { const td = document.createElement('td'); row.appendChild(td); return td; }
        const warehouse = document.createElement('select');
        warehouse.className = 'form-select form-select-sm bg-dark text-white border-secondary';
        warehouse.setAttribute('aria-label', '发货仓库');
        warehouse.appendChild(new Option('所有仓库（区域通用）', ''));
        warehouses.forEach(w => warehouse.appendChild(new Option(w.name + (w.is_active ? '' : '（停用）'), w.id)));
        warehouse.value = rule.warehouse_id == null ? '' : String(rule.warehouse_id);
        // A deleted warehouse must not silently become an all-warehouse rule.
        if (rule.warehouse_id != null && warehouse.value !== String(rule.warehouse_id)) {
            warehouse.appendChild(new Option('仓库已删除：' + rule.warehouse_id, rule.warehouse_id));
            warehouse.value = String(rule.warehouse_id);
        }
        cell().appendChild(warehouse);
        function input(label, value, type, list) {
            const el = document.createElement('input');
            el.type = type; el.value = value;
            el.className = 'form-control form-control-sm bg-dark text-white border-secondary';
            el.style.minWidth = '90px'; el.setAttribute('aria-label', label); el.required = true;
            if (list) el.setAttribute('list', list);
            if (type === 'number') { el.min = '0'; el.max = '1000000'; el.step = '0.01'; }
            cell().appendChild(el); return el;
        }
        const country = input('收货国家代码', rule.destination_country || '', 'text', 'returnLossCountries');
        country.placeholder = 'CZ / PL'; country.maxLength = 2; country.pattern = '[A-Za-z]{2}';
        const currency = input('币种', rule.currency || '', 'text', 'returnLossCurrencies');
        currency.placeholder = 'CZK'; currency.maxLength = 3; currency.pattern = '[A-Za-z]{3}';
        const outbound = input('正向运费', rule.outbound_amount ?? '', 'number');
        const inbound = input('逆向运费', rule.return_amount ?? '', 'number');
        const total = cell(); total.className = 'text-warning text-nowrap';
        function updateTotal() {
            const values = [outbound.value, inbound.value];
            total.textContent = values.every(v => v !== '' && Number.isFinite(Number(v)) && Number(v) >= 0)
                ? (Number(values[0]) + Number(values[1])).toFixed(2) + ' ' + currency.value.toUpperCase() : '—';
        }
        [outbound, inbound, currency].forEach(el => el.addEventListener('input', updateTotal));
        const remove = document.createElement('button'); remove.type = 'button';
        remove.className = 'btn btn-outline-danger btn-sm'; remove.textContent = '移除';
        remove.addEventListener('click', () => { row.remove(); say('规则已修改，点击保存后生效'); });
        cell().appendChild(remove);
        row.readRule = () => ({warehouse_id: warehouse.value ? Number(warehouse.value) : null,
            destination_country: country.value.trim().toUpperCase(), currency: currency.value.trim().toUpperCase(),
            outbound_amount: outbound.value, return_amount: inbound.value});
        body.appendChild(row); updateTotal();
    }
    async function load() {
        if (busy) return;
        lock(true);
        try {
            const response = await fetch('/api/settings/return-shipping-loss');
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || '加载失败');
            version = data.version; warehouses = data.warehouses;
            body.replaceChildren(); data.rules.forEach(addRow);
            say(data.rules.length ? '已加载 ' + data.rules.length + ' 条规则' : '尚未配置规则，当前沿用订单运费');
        } catch (error) { say(error.message, true); }
        finally { lock(false); }
    }
    add.addEventListener('click', () => { addRow(); say('填写国家代码、币种和运费后保存'); });
    reload.addEventListener('click', load);
    save.addEventListener('click', async () => {
        if (busy || version === null) return;
        for (const input of body.querySelectorAll('input')) {
            if (!input.reportValidity()) return;
        }
        const rules = Array.from(body.children, row => row.readRule());
        lock(true);
        try {
            const response = await fetch('/api/settings/return-shipping-loss', {method: 'POST',
                headers: {'Content-Type': 'application/json'}, body: JSON.stringify({rules, version})});
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || '保存失败');
            version = data.version;
            say('已保存。后续退件按新规则计算，历史金额保持不变。');
        } catch (error) { say(error.message, true); }
        finally { lock(false); }
    });
    load();
})();
