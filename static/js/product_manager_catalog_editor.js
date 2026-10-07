(function (global) {
    'use strict';
    const FIELDS = ['manage_stock', 'stock_quantity', 'stock_status', 'regular_price', 'sale_price'];
    const STOCK_FIELDS = ['manage_stock', 'stock_quantity', 'stock_status'];
    const STOCK_LABELS = {instock: '有货', outofstock: '售完', onbackorder: '可预订'};
    const str = value => value === null || value === undefined ? '' : String(value);
    const escape = value => str(value).replace(/[&<>"']/g, character => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[character]));
    const copy = value => value === undefined ? undefined : JSON.parse(JSON.stringify(value));
    function stable(value) {
        if (Array.isArray(value)) return value.map(stable);
        if (value && typeof value === 'object') return Object.fromEntries(Object.keys(value).sort().map(key => [key, stable(value[key])]));
        return value;
    }
    const sameIdentity = (left, right) => JSON.stringify(stable(left)) === JSON.stringify(stable(right));
    const money = value => {
        const text = str(value).trim();
        if (!text) return '';
        return text.replace(/^0+(?=\d)/, '').replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
    };
    function sameField(field, left, right) {
        if (field === 'stock_quantity') return left === null || left === undefined || left === '' ? right === null || right === undefined || right === '' : Number(left) === Number(right);
        if (field === 'regular_price' || field === 'sale_price') return money(left) === money(right);
        return left === right;
    }
    const rowKey = row => `${row.site_id}:${row.product_id}:${row.variation_id || 0}`;
    const siteName = row => str(row.site && (row.site.name || row.site.url) || ('站点 ' + row.site_id));
    const beforeRow = row => copy(row.edit_before || Object.fromEntries(FIELDS.map(field => [field, row[field]])));
    const inherited = row => beforeRow(row).manage_stock === 'parent';

    function validChanges(changes, baseline) {
        const output = {};
        Object.entries(changes).forEach(([field, value]) => {
            if (!FIELDS.includes(field)) throw new Error('不支持的修改字段。');
            if (field === 'manage_stock') {
                if (typeof value !== 'boolean') throw new Error('请选择是否启用独立库存管理。');
                output[field] = value;
            } else if (field === 'stock_quantity') {
                if (!/^\d+$/.test(str(value).trim()) || !Number.isSafeInteger(Number(value)) || Number(value) > 2147483647) throw new Error('库存数量须为 0 至 2147483647 之间的整数。');
                output[field] = Number(value);
            } else if (field === 'stock_status') {
                if (!Object.prototype.hasOwnProperty.call(STOCK_LABELS, value)) throw new Error('请选择有效库存状态。');
                output[field] = value;
            } else {
                const text = str(value).trim();
                if (text && (text.length > 64 || !/^\d{1,13}(?:\.\d{1,8})?$/.test(text))) throw new Error('价格须为非负数字，最多 13 位整数和 8 位小数。');
                output[field] = field === 'sale_price' && text && money(text) === '0' ? '' : text;
            }
        });
        if (baseline.manage_stock === 'parent' && STOCK_FIELDS.some(field => field in output) && (output.manage_stock !== true || !('stock_quantity' in output))) {
            throw new Error('此口味沿用整款商品库存，请明确启用独立库存并填写该口味数量；不能复制整款库存或修改未选口味。');
        }
        if ('stock_quantity' in output && !('manage_stock' in output) && baseline.manage_stock === true) output.manage_stock = true;
        if ('stock_status' in output && !('manage_stock' in output) && baseline.manage_stock === false) output.manage_stock = false;
        if (output.manage_stock === true && !('stock_quantity' in output)) {
            if (baseline.stock_quantity === null || baseline.stock_quantity === undefined) throw new Error('启用独立库存管理时，请填写库存数量。');
            output.stock_quantity = Number(baseline.stock_quantity);
        }
        if (output.manage_stock === false && !('stock_status' in output)) {
            if (!Object.prototype.hasOwnProperty.call(STOCK_LABELS, baseline.stock_status)) throw new Error('关闭库存管理时，请选择库存状态。');
            output.stock_status = baseline.stock_status;
        }
        return output;
    }

    async function requestJson(fetcher, url, options) {
        const aborter = new AbortController();
        let timedOut = false;
        const timer = setTimeout(() => { timedOut = true; aborter.abort(); }, 120000);
        try {
            let response;
            try { response = await fetcher(url, Object.assign({method: 'GET', headers: {Accept: 'application/json'}}, options || {}, {signal: aborter.signal})); }
            catch (error) { throw Object.assign(new Error(timedOut ? '请求超时，提交结果可能需要只读核对，请勿直接重复提交。' : '网络请求未完成：' + (error.message || '连接异常')), {httpStatus: 0}); }
            const type = response.headers && response.headers.get('content-type') || '';
            if (!/\bapplication\/(?:[\w.+-]+\+)?json\b/i.test(type)) throw Object.assign(new Error(`接口返回非 JSON 内容（HTTP ${response.status}），请核对登录或网关状态。`), {httpStatus: response.status});
            let data;
            try { data = typeof response.json === 'function' ? await response.json() : JSON.parse(await response.text()); }
            catch (_) { throw Object.assign(new Error(timedOut ? '响应读取超时，提交结果需要核对，请勿重复提交。' : `接口返回内容无法解析（HTTP ${response.status}）。`), {httpStatus: response.status}); }
            if (!response.ok || !data || data.error || data.success === false) throw Object.assign(new Error(str(data && (data.error || data.message)) || `请求失败（HTTP ${response.status}）`),
                {httpStatus: response.status, code: data && data.code, writeStarted: data && data.write_started, verificationStatus: data && data.verification_status, expectedStockBridge: data && data.expected_stock_bridge});
            if (typeof data !== 'object' || Array.isArray(data)) throw new Error('接口 JSON 格式无效，请核对结果。');
            return data;
        } finally { clearTimeout(timer); }
    }

    function createEditor(options) {
        const catalog = options.catalog, fetcher = options.fetch;
        const selected = new Set(), drafts = new Map(), results = new Map(), intents = new Map(), jobs = new Map();
        let config = null, configError = '', busy = false, action = '', tableVersion = 0, lastRows = null, rowsByKey = new Map();
        let pollTimer = null, cloneSkippedSelf = 0, stopRequested = false, actionKind = '', cloneSubmissionLock = null;
        const newId = () => options.uuid ? options.uuid() : global.crypto && global.crypto.randomUUID ? global.crypto.randomUUID() : `catalog-${Date.now()}-${Math.random().toString(16).slice(2)}`;
        function rowMap() {
            const rows = catalog.getSnapshot().rows;
            if (rows !== lastRows) { lastRows = rows; rowsByKey = new Map(rows.map(row => [rowKey(row), row])); }
            return rowsByKey;
        }
        function capability(row, name) { return Boolean((name !== 'can_edit' || ['simple', 'variation'].includes(row.type)) && config && config.sites.some(site => str(site.id) === str(row.site_id) && site[name] === true)); }
        function notify(table) { if (table) tableVersion += 1; if (options.onChange) options.onChange(); }
        function getSnapshot() {
            const snapshot = catalog.getSnapshot(), available = snapshot.rows.filter(row => selected.has(rowKey(row))), visible = snapshot.filtered.filter(row => selected.has(rowKey(row)));
            return {ready: !!config, configError, config, busy, action, tableVersion, stopRequested, canStop: busy && actionKind === 'edit' && !stopRequested, selected: [...selected], selectedCount: selected.size, selectedAvailable: available.length,
                selectedVisible: visible.length, selectedHidden: selected.size - visible.length,
                drafts: [...drafts.entries()].map(([key, draft]) => ({key, values: copy(draft.values)})), draftCount: drafts.size,
                results: [...results.values()].map(copy), jobs: [...jobs.values()].map(copy), cloneSkippedSelf, cloneSubmissionLock: copy(cloneSubmissionLock)};
        }
        const configPromise = requestJson(fetcher, '/api/product-manager/catalog-edit-config').then(data => {
            if (!Array.isArray(data.sites) || !Array.isArray(data.clone_targets) || !data.csrf_token) throw new Error('编辑权限信息不完整，请刷新页面。');
            config = data; notify(true); return data;
        }).catch(error => { configError = error.message; notify(true); return null; });
        function baseline(row) {
            if (!row.edit_identity || !row.edit_before) throw new Error('此条结果缺少编辑核对信息，请重新加载后修改。');
            return {identity: copy(row.edit_identity), before: copy(row.edit_before)};
        }
        function setDraft(key, field, value) {
            if (busy) return;
            if (results.get(key)?.status === 'unconfirmed') throw new Error('此行提交结果尚未确认，请先只读核对或重新加载，勿重复修改。');
            const row = rowMap().get(key);
            if (!row || !capability(row, 'can_edit') || !FIELDS.includes(field)) return;
            let draft = drafts.get(key);
            if (!draft) draft = {baseline: baseline(row), values: {}};
            if (field === 'manage_stock' && value === false && draft.baseline.before.manage_stock === 'parent') {
                delete draft.values.manage_stock; delete draft.values.stock_quantity; delete draft.values.stock_status;
            } else if (field === 'stock_quantity' && draft.baseline.before.manage_stock === 'parent' && draft.values.manage_stock === true) {
                // An explicitly entered independent quantity remains meaningful even
                // when its number happens to equal the shared parent total.
                draft.values.stock_quantity = value;
            } else if (sameField(field, value, draft.baseline.before[field]) || field === 'sale_price' && money(value) === '0' && !draft.baseline.before.sale_price) delete draft.values[field];
            else draft.values[field] = value;
            if (field === 'manage_stock' && value === false) delete draft.values.stock_quantity;
            if (Object.keys(draft.values).length) drafts.set(key, draft); else drafts.delete(key);
            results.delete(key); notify(false);
        }
        function draftChanges(key) {
            const draft = drafts.get(key);
            return draft ? validChanges(draft.values, draft.baseline.before) : {};
        }
        function select(keys, mode) {
            if (busy) return;
            keys.forEach(key => {
                const row = rowMap().get(key);
                if (!row || !(capability(row, 'can_edit') || capability(row, 'can_clone'))) return;
                if (mode === false) selected.delete(key); else selected.add(key);
            });
            notify(true);
        }
        function clearSelection() { if (!busy) { selected.clear(); notify(true); } }
        function discard(key) { if (busy) return; drafts.delete(key); results.delete(key); notify(true); }
        function itemUrl(row) {
            return '/api/product-manager/catalog-item?' + new URLSearchParams({site_id: row.site_id, product_id: row.product_id, variation_id: row.variation_id || 0});
        }
        async function execute(row, patch, expected, batchId) {
            const key = rowKey(row);
            let sent = false;
            try {
                const changes = validChanges(patch, expected.before);
                if (!Object.keys(changes).length) return {key, status: 'unchanged'};
                intents.set(key, {row: copy(row), changes: copy(changes), baseline: copy(expected)});
                const fresh = await requestJson(fetcher, itemUrl(row));
                if (stopRequested) return {key, status: 'stopped', error: '已停止，本行尚未提交修改。'};
                if (!sameIdentity(expected.identity, fresh.identity) || Object.keys(changes).some(field => !sameField(field, expected.before[field], fresh.before[field]))) {
                    throw Object.assign(new Error('站点商品身份或拟修改字段已变化，请重新加载核对；本次未提交修改。'), {httpStatus: 409, code: 'stale_loaded_evidence'});
                }
                const intent = intents.get(key);
                intent.stockMutation = STOCK_FIELDS.some(field => field in changes);
                intent.bridgePresent = fresh.stock_bridge && typeof fresh.stock_bridge.present === 'boolean' ? fresh.stock_bridge.present : null;
                if (intent.stockMutation && intent.bridgePresent === true) {
                    const desired = Object.assign({}, fresh.before, changes), quantity = Number(desired.stock_quantity);
                    intent.expectedBridge = {wcms_stock_manage: desired.manage_stock ? 'yes' : 'no', wcms_stock_qty: Number.isFinite(quantity) ? Math.max(0, Math.trunc(quantity)) : 0, wcms_stock_status: desired.stock_status || 'instock'};
                }
                sent = true;
                const data = await requestJson(fetcher, itemUrl(row), {method: 'PUT', headers: {'Content-Type': 'application/json', Accept: 'application/json', 'X-PM-CSRF': fresh.csrf_token || config.csrf_token},
                    body: JSON.stringify({site_id: Number(row.site_id), product_id: row.product_id, variation_id: row.variation_id || 0, expected_identity: fresh.identity, expected_before: fresh.before, changes, batch_id: batchId})});
                if (data.success !== true || data.verification?.status !== 'verified' || !data.item || rowKey(data.item) !== key || !sameIdentity(expected.identity, data.item.edit_identity) ||
                    Object.entries(changes).some(([field, value]) => !sameField(field, value, data.item.edit_before?.[field]))) throw new Error('服务器未返回此站点商品符合修改目标的已核验结果，请先核对状态。');
                catalog.updateRows([data.item]); drafts.delete(key); selected.delete(key); intents.delete(key);
                return {key, status: 'verified', site_id: row.site_id, product_id: row.product_id, variation_id: row.variation_id || 0};
            } catch (error) {
                if (error.expectedStockBridge && intents.has(key)) intents.get(key).expectedBridge = copy(error.expectedStockBridge);
                // HTTP 409 can also mean the remote PUT happened but final read-back
                // failed. Only explicit pre-write evidence permits automatic retry.
                const writeUncertain = error.writeStarted === true || error.verificationStatus === 'unconfirmed';
                const definitelyRejected = !sent || !writeUncertain && (error.writeStarted === false || ['AUTH_REQUIRED', 'CSRF_REJECTED'].includes(error.code));
                return {key, status: definitelyRejected ? 'failed' : 'unconfirmed', error: error.message || str(error), code: error.code || '', site_id: row.site_id, product_id: row.product_id, variation_id: row.variation_id || 0};
            }
        }
        async function runItems(items, label) {
            await configPromise;
            if (!config) throw new Error(configError || '编辑权限尚未读取。');
            if (busy || catalog.getSnapshot().busy) throw new Error('请等待当前加载或操作结束。');
            if (!items.length) throw new Error('请先选择可修改的产品或填写行修改。');
            if (items.some(item => results.get(rowKey(item.row))?.status === 'unconfirmed')) throw new Error('有提交结果未确认的行，请先只读核对，勿重复提交。');
            items.forEach(item => results.delete(rowKey(item.row)));
            busy = true; actionKind = 'edit'; stopRequested = false; action = label + ' 0 / ' + items.length; notify(true);
            const batchId = newId(), workers = items.some(item => STOCK_FIELDS.some(field => field in item.changes)) ? 1 : 2;
            let cursor = 0, completed = 0;
            async function worker() {
                while (!stopRequested && cursor < items.length) {
                    const item = items[cursor++];
                    let result;
                    if (!capability(item.row, 'can_edit')) result = {key: rowKey(item.row), status: 'failed', error: '当前账号没有该站点的修改权限。'};
                    else result = await execute(item.row, item.changes, item.baseline, batchId);
                    results.set(result.key, result); completed += 1; action = `${label} ${completed} / ${items.length}`; notify(true);
                }
            }
            try { await Promise.all(Array.from({length: Math.min(workers, items.length)}, worker)); }
            finally {
                items.forEach(item => { if (!results.has(rowKey(item.row))) results.set(rowKey(item.row), {key: rowKey(item.row), status: 'stopped', error: '已停止，本行尚未提交修改。'}); });
                busy = false; actionKind = ''; action = `${label}${stopRequested ? '已停止' : '结束'}：已核验 ${items.filter(item => results.get(rowKey(item.row))?.status === 'verified').length} / ${items.length}${stopRequested ? '，已发送的修改已继续核对，未提交项保留选择 / 修改' : ''}`; notify(true);
            }
            return items.map(item => results.get(rowKey(item.row)));
        }
        function saveDrafts(keys) {
            const chosen = keys || [...drafts.keys()], items = [];
            chosen.forEach(key => {
                if (results.get(key)?.status === 'unconfirmed') throw new Error('有提交结果未确认的行，请先只读核对，勿重复保存。');
                const row = rowMap().get(key), draft = drafts.get(key);
                if (row && draft) items.push({row, changes: draftChanges(key), baseline: draft.baseline});
            });
            return runItems(items, '保存修改');
        }
        function bulk(operation, values) {
            const rows = [...rowMap().values()].filter(row => selected.has(rowKey(row)));
            if (rows.some(row => results.get(rowKey(row))?.status === 'unconfirmed')) throw new Error('所选包含提交结果未确认的行，请先只读核对，勿重复修改。');
            if (rows.some(row => drafts.has(rowKey(row)))) throw new Error('所选产品有未保存的行修改，请先保存或撤销后再批量操作。');
            let changes;
            values = values || {};
            if (operation === 'price') {
                changes = {};
                if (str(values.regular_price).trim()) changes.regular_price = str(values.regular_price).trim();
                if (str(values.sale_price).trim()) changes.sale_price = money(values.sale_price) === '0' ? '' : str(values.sale_price).trim();
                if (!Object.keys(changes).length) throw new Error('至少填写一个价格；留空字段保持不变。');
            } else if (operation === 'stock' || operation === 'restore') changes = {manage_stock: true, stock_quantity: values.stock_quantity};
            else if (operation === 'stock_status') changes = {manage_stock: false, stock_status: values.stock_status};
            else if (operation === 'soldout') changes = {manage_stock: false, stock_status: 'outofstock'};
            else if (operation === 'clear_sale') changes = {sale_price: ''};
            else throw new Error('不支持的批量操作。');
            if (rows.some(inherited) && ['stock', 'restore'].includes(operation) && values.confirm_independent !== true) throw new Error('包含沿用父库存的口味，请确认启用各口味的独立库存管理。');
            const items = rows.map(row => ({row, changes: validChanges(changes, beforeRow(row)), baseline: baseline(row)}));
            return runItems(items, '批量操作');
        }
        async function verify(key) {
            const intent = intents.get(key);
            if (!intent || busy) return;
            busy = true; notify(true);
            try {
                const fresh = await requestJson(fetcher, itemUrl(intent.row));
                if (!fresh.item || rowKey(fresh.item) !== key || !sameIdentity(intent.baseline.identity, fresh.identity) || Object.entries(intent.changes).some(([field, value]) => !sameField(field, value, fresh.before[field]))) throw new Error('回读结果仍未符合修改目标，请重新加载人工核对，勿重复提交。');
                if (intent.stockMutation) {
                    if (intent.expectedBridge) {
                        // Woo's quantity threshold determines final stock status;
                        // changed core fields were checked against the exact intent.
                        const finalQuantity = Number(fresh.before.stock_quantity), expectedBridge = {wcms_stock_manage: fresh.before.manage_stock === true ? 'yes' : 'no',
                            wcms_stock_qty: Number.isFinite(finalQuantity) ? Math.max(0, Math.trunc(finalQuantity)) : 0, wcms_stock_status: fresh.before.stock_status || 'instock'};
                        if (fresh.stock_bridge?.present !== true || fresh.stock_bridge.valid !== true || Object.entries(expectedBridge).some(([field, value]) => str(fresh.stock_bridge.values?.[field]) !== str(value))) throw new Error('商品核心库存字段已回读，但 WCMS 库存仍未达到目标，请继续人工核对，勿重复提交。');
                    } else if (intent.bridgePresent !== false || fresh.stock_bridge?.present !== false || fresh.stock_bridge.valid !== true) throw new Error('缺少库存联动核对证据，请重新加载并人工核对，不能将此行标记为已完成。');
                }
                catalog.updateRows([fresh.item]); drafts.delete(key); selected.delete(key); intents.delete(key);
                results.set(key, {key, status: 'verified', verified_by_read_only_check: true});
            } finally { busy = false; notify(true); }
        }
        function retryFailed() {
            const items = [...results.values()].filter(result => result.status === 'failed' && intents.has(result.key)).map(result => {
                const intent = intents.get(result.key); return {row: intent.row, changes: intent.changes, baseline: intent.baseline};
            });
            return runItems(items, '重试失败项');
        }
        async function cloneSelected(cloneOptions) {
            await configPromise;
            if (!config || busy || catalog.getSnapshot().busy) throw new Error('请等待加载或当前操作结束。');
            if (cloneSubmissionLock) throw new Error('此前克隆提交结果未完全确认，已阻止重复提交。请核对已知后台任务；丢失任务编号的批次请联系管理员核对任务及目标商品。');
            const targetId = Number(cloneOptions.target_site_id);
            if (!config.clone_targets.some(site => Number(site.id) === targetId)) throw new Error('请选择有克隆权限的目标站点。');
            const groups = new Map();
            [...rowMap().values()].filter(row => selected.has(rowKey(row))).forEach(row => {
                if (!capability(row, 'can_clone')) throw new Error('所选源站没有克隆权限。');
                if (!groups.has(str(row.site_id))) groups.set(str(row.site_id), new Set());
                groups.get(str(row.site_id)).add(row.product_id);
            });
            cloneSkippedSelf = groups.get(str(targetId))?.size || 0;
            groups.delete(str(targetId));
            if (!groups.size) throw new Error('目标站自身的商品无需克隆，请选择其他源站产品。');
            const created = []; busy = true; actionKind = 'clone'; action = '正在提交后台克隆任务…'; notify(true);
            try {
                for (const [source, ids] of groups) {
                    const productIds = [...ids];
                    for (let offset = 0; offset < productIds.length; offset += 50) {
                        const payload = {source_site_id: Number(source), target_site_id: targetId, product_ids: productIds.slice(offset, offset + 50), include_variations: cloneOptions.include_variations !== false,
                            include_images: cloneOptions.include_images !== false, status_on_target: cloneOptions.collision_mode === 'clone_as_new' ? 'draft' : cloneOptions.status_on_target || 'draft', collision_mode: cloneOptions.collision_mode || 'skip_existing'};
                        const response = await requestJson(fetcher, '/api/product-manager/catalog-clone', {method: 'POST', headers: {'Content-Type': 'application/json', Accept: 'application/json', 'X-PM-CSRF': config.csrf_token}, body: JSON.stringify(payload)});
                        if (!response.job_id) throw new Error('克隆提交未确认，勿重复提交，请核对后台任务。');
                        const job = {job_id: str(response.job_id), source_site_id: source, target_site_id: str(targetId), total_count: response.total_count || payload.product_ids.length, completed_count: 0, status: response.status || 'queued', terminal: false};
                        jobs.set(job.job_id, job); created.push(copy(job)); persistJobs(); notify(false);
                    }
                }
                action = `已提交 ${created.length} 个后台克隆任务${cloneSkippedSelf ? `；目标站自身 ${cloneSkippedSelf} 款未克隆` : ''}。`;
            } catch (error) {
                if (created.length || error.writeStarted !== false) {
                    cloneSubmissionLock = {reason: error.message || '克隆提交未确认', known_job_ids: created.map(job => job.job_id), at: new Date().toISOString()};
                    persistCloneLock();
                }
                action = '克隆提交未完全确认：' + error.message + '。已知任务继续执行；未知任务编号请联系管理员核对，勿重复提交。'; throw error;
            }
            finally { busy = false; actionKind = ''; notify(true); schedulePoll(); }
            return created;
        }
        function persistJobs() {
            if (!options.storage) return;
            try { options.storage.setItem('productManagerCatalogCloneJobs', JSON.stringify([...jobs.values()].filter(job => !job.terminal))); } catch (_) { /* Browser storage may be unavailable. */ }
        }
        function persistCloneLock() {
            if (!options.storage) return;
            try { if (cloneSubmissionLock) options.storage.setItem('productManagerCatalogCloneSubmissionLock', JSON.stringify(cloneSubmissionLock)); else options.storage.removeItem('productManagerCatalogCloneSubmissionLock'); } catch (_) { /* Storage is optional. */ }
        }
        async function pollJobs() {
            const pending = [...jobs.values()].filter(job => !job.terminal);
            let cursor = 0;
            async function worker() {
                while (cursor < pending.length) {
                    const job = pending[cursor++];
                    try {
                        const response = await requestJson(fetcher, '/api/product-manager/clone-jobs/' + encodeURIComponent(job.job_id));
                        const details = response.results && typeof response.results === 'object' ? response.results : {};
                        jobs.set(job.job_id, Object.assign({}, job, {status: response.status, terminal: response.terminal === true, completed_count: Number(response.completed_count || 0), total_count: Number(response.total_count || job.total_count),
                            created_count: Number(response.created_count || 0), skipped_count: Number(response.skipped_count || 0), failed_count: Number(response.failed_count || 0), last_error: str(response.last_error), poll_error: '',
                            failed_details: (Array.isArray(details.failed) ? details.failed : []).slice(0, 50).map(item => ({product_id: item.product_id, target_id: item.target_id, partial_clone: item.partial_clone === true, error: str(item.error), warnings: Array.isArray(item.warnings) ? item.warnings.map(str) : []})),
                            success_warnings: (Array.isArray(details.success) ? details.success : []).filter(item => Array.isArray(item.warnings) && item.warnings.length).slice(0, 20).map(item => ({product_id: item.source_id || item.product_id, target_id: item.target_id, warnings: item.warnings.map(str)}))}));
                    } catch (error) { job.poll_error = error.message; }
                    notify(false);
                }
            }
            await Promise.all(Array.from({length: Math.min(2, pending.length)}, worker)); persistJobs(); schedulePoll();
            return [...jobs.values()].map(copy);
        }
        function schedulePoll() {
            if (options.autoPoll === false || pollTimer !== null || ![...jobs.values()].some(job => !job.terminal)) return;
            pollTimer = setTimeout(() => { pollTimer = null; pollJobs(); }, options.pollInterval || 2000);
        }
        if (options.storage) {
            try { JSON.parse(options.storage.getItem('productManagerCatalogCloneJobs') || '[]').forEach(job => { if (job.job_id) jobs.set(str(job.job_id), job); }); } catch (_) { /* Ignore malformed local progress data. */ }
            try { cloneSubmissionLock = JSON.parse(options.storage.getItem('productManagerCatalogCloneSubmissionLock') || 'null'); } catch (_) { cloneSubmissionLock = {reason: '本页克隆提交记录无法读取，请先人工核对。'}; }
            schedulePoll();
        }
        return {ready: () => configPromise, getSnapshot, select, clearSelection, setDraft, saveDrafts, bulk, cloneSelected, pollJobs, verify, retryFailed, discard,
            stopWrites() { if (busy && actionKind === 'edit') { stopRequested = true; action = '正在停止后续提交；已发送的修改继续回读核对…'; notify(false); } },
            acknowledgeCloneSubmission() {
                if (busy || !cloneSubmissionLock) return false;
                if (!(options.confirm || global.confirm)('请先核对已知任务的结果。丢失任务编号的批次须联系管理员核对后台任务与目标商品。你已完成核对，并确认新的克隆不会重复已提交商品吗？')) return false;
                cloneSubmissionLock = null; persistCloneLock(); notify(true); return true;
            },
            getDraft: key => drafts.get(key)?.values || {}, getResult: key => results.get(key), capability,
            canReload() {
                if (busy) return false;
                if (drafts.size && !(options.confirm || global.confirm)(`还有 ${drafts.size} 行未保存。重新加载会清除这些修改，是否继续？`)) return false;
                drafts.clear(); selected.clear(); results.clear(); intents.clear(); notify(true); return true;
            }};
    }

    function createUI(options) {
        const document = options.document, root = options.root, catalog = options.catalog;
        const $ = id => document.getElementById(id);
        let localError = '', modalOperation = '', modalKeys = [], modalBusy = false;
        let browserStorage;
        try { browserStorage = global.sessionStorage; } catch (_) { /* Storage is optional. */ }
        const editor = createEditor({catalog, fetch: options.fetch, storage: browserStorage, confirm: message => global.confirm(message),
            onChange() { if (options.onChange) options.onChange(); }});
        const stockStatus = row => STOCK_LABELS[row.stock_status] || str(row.stock_status) || '未知';
        const rowByKey = key => catalog.getSnapshot().rows.find(row => rowKey(row) === key);
        const safeLink = value => { try { const url = new URL(value); return ['http:', 'https:'].includes(url.protocol) ? url.href : ''; } catch (_) { return ''; } };
        function valuesFor(row) {
            const original = beforeRow(row), draft = editor.getDraft(rowKey(row));
            const values = Object.assign({}, original, draft);
            if (original.manage_stock === 'parent' && draft.manage_stock === true && !('stock_quantity' in draft)) values.stock_quantity = '';
            return values;
        }
        function resultLabel(key) {
            const result = editor.getResult(key);
            if (result?.status === 'verified') return '已保存并回读核验';
            if (result?.status === 'unconfirmed') return '提交结果未确认，请先只读核对，勿重复保存。' + (result.error ? ' ' + result.error : '');
            if (result?.status === 'failed') return '未完成：' + result.error;
            if (result?.status === 'stopped') return '已停止，此行尚未提交修改';
            return Object.keys(editor.getDraft(key)).length ? '有未保存修改' : '';
        }
        function tableHtml(snapshot) {
            const state = editor.getSnapshot(), selected = new Set(state.selected), disabled = state.busy || snapshot.busy;
            const rows = snapshot.pageRows.map(row => {
                const key = rowKey(row), values = valuesFor(row), editable = editor.capability(row, 'can_edit') && !!row.edit_identity && !!row.edit_before;
                const selectable = editable || editor.capability(row, 'can_clone');
                const inheritedStock = inherited(row), independent = values.manage_stock === true;
                const link = safeLink(row.permalink), dirty = Object.keys(editor.getDraft(key)).length > 0, result = editor.getResult(key);
                const controlsOff = disabled || !editable || result?.status === 'unconfirmed';
                const rowClass = result?.status === 'verified' ? 'pm-row-success' : ['failed', 'unconfirmed'].includes(result?.status) ? 'pm-row-error' : dirty ? 'pm-row-modified' : '';
                const input = (field, value, extra = '') => `<input data-edit-field="${field}" aria-label="${escape(row.product_name)} ${field === 'regular_price' ? '原价' : field === 'sale_price' ? '优惠价' : '库存数量'}" value="${escape(value)}" class="form-control form-control-sm bg-dark text-white border-secondary ${field === 'stock_quantity' ? 'pm-stock-input' : 'pm-price-input'}" ${extra} ${controlsOff ? 'disabled' : ''}>`;
                return `<tr data-catalog-key="${escape(key)}" class="${rowClass}"><td><input type="checkbox" class="form-check-input" data-edit-select="${escape(key)}" aria-label="选择 ${escape(siteName(row))} ${escape(row.product_name)} ${escape((row.flavors || []).join(' / '))}" ${selected.has(key) ? 'checked' : ''} ${disabled || !selectable ? 'disabled' : ''}></td>` +
                    `<td>${escape(siteName(row))}<div class="text-muted small">${escape(row.site && row.site.manager || '未设置负责人')}</div></td><td>${escape((row.brands || []).join(' / ') || '未识别')}</td>` +
                    `<td class="pm-name">${link ? `<a href="${escape(link)}" target="_blank" rel="noopener noreferrer">${escape(row.product_name)}</a>` : escape(row.product_name)}<div class="small text-muted">${escape(({publish: '已发布', draft: '草稿', pending: '待审核', private: '私密', future: '定时发布'})[row.status] || row.status || '未知')} · ${row.variation_id ? '变体 ' + row.variation_id : '产品 ' + row.product_id}</div></td>` +
                    `<td>${escape((row.flavors || []).join(' / ') || '未识别口味')}${row.flavor_scope === 'any' ? '<div class="small text-info">任意口味（共用此变体）</div>' : ''}<div class="pm-sku">SKU：${escape(row.sku || '未提供')}</div></td>` +
                    `<td style="min-width:190px"><div class="small ${row.stock_status === 'outofstock' ? 'text-danger' : 'text-info'}">${escape(stockStatus(row))}${inheritedStock ? ' · 沿用整款库存' : ''}</div><label class="form-check small mb-1"><input type="checkbox" class="form-check-input" data-edit-field="manage_stock" ${independent ? 'checked' : ''} ${controlsOff ? 'disabled' : ''}>${inheritedStock ? '改为此口味独立库存' : '管理库存数量'}</label>` +
                    `<div class="pm-stock-line"><label class="pm-stock-line-label">数量</label>${input('stock_quantity', values.stock_quantity, 'type="number" min="0" max="2147483647" step="1"' + (!independent ? ' disabled' : ''))}</div>` +
                    `<div class="pm-stock-line mt-1"><label class="pm-stock-line-label">状态</label><select data-edit-field="stock_status" aria-label="${escape(row.product_name)} 库存状态" class="form-select form-select-sm pm-status-input" ${controlsOff || independent || inheritedStock ? 'disabled' : ''}>${Object.entries(STOCK_LABELS).map(([value, label]) => `<option value="${value}" ${values.stock_status === value ? 'selected' : ''}>${label}</option>`).join('')}</select></div>` +
                    `${inheritedStock ? '<div class="small text-muted mt-1" data-edit-inherited-note>独立数量须另行填写，不能直接关闭共享库存。</div>' : ''}</td>` +
                    `<td><label class="small text-muted">原价</label>${input('regular_price', values.regular_price, 'type="text" inputmode="decimal" maxlength="64"')}<label class="small text-muted mt-1">优惠价</label>${input('sale_price', values.sale_price, 'type="text" inputmode="decimal" maxlength="64"')}<div class="small text-muted">留空 / 0 清除优惠价</div></td>` +
                    `<td style="min-width:170px"><div class="d-flex gap-1 flex-wrap"><button type="button" class="btn btn-sm btn-primary" data-edit-save="${escape(key)}" ${disabled || !dirty || !editable || result?.status === 'unconfirmed' ? 'disabled' : ''}>保存</button><button type="button" class="btn btn-sm btn-outline-secondary" data-edit-discard="${escape(key)}" ${disabled || !dirty || result?.status === 'unconfirmed' ? 'disabled' : ''}>撤销</button>${result?.status === 'unconfirmed' ? `<button type="button" class="btn btn-sm btn-outline-warning" data-edit-verify="${escape(key)}" ${disabled ? 'disabled' : ''}>只读核对</button>` : ''}<button type="button" class="btn btn-sm btn-outline-light" data-catalog-open="${escape(key)}">单站管理</button></div><div class="small mt-1 ${['failed', 'unconfirmed'].includes(result?.status) ? 'text-warning' : 'text-muted'}" data-edit-feedback>${escape(resultLabel(key))}</div>${!editable ? '<div class="small text-muted">跨站编辑仅支持有修改权限的简单商品和明确变体</div>' : ''}</td></tr>`;
            }).join('');
            return '<div class="table-responsive"><table class="table table-dark table-hover pm-table mb-0"><thead><tr><th>选择</th><th>站点 / 负责人</th><th>品牌</th><th>产品 / 发布状态</th><th>口味 / SKU</th><th>库存</th><th>价格（站点原币）</th><th>操作</th></tr></thead><tbody>' + rows + '</tbody></table></div>' +
                `<div class="d-flex justify-content-between align-items-center p-3 gap-2"><span class="text-muted small">第 ${snapshot.page} / ${snapshot.pages} 页，每页最多 100 条；选择和修改跨页保留</span><div class="d-flex gap-2"><button type="button" class="btn btn-sm btn-outline-light" data-catalog-page="${snapshot.page - 1}" ${snapshot.page === 1 ? 'disabled' : ''}>上一页</button><button type="button" class="btn btn-sm btn-outline-light" data-catalog-page="${snapshot.page + 1}" ${snapshot.page === snapshot.pages ? 'disabled' : ''}>下一页</button></div></div>`;
        }
        function render(snapshot) {
            const state = editor.getSnapshot(), unavailable = !state.ready || state.busy || snapshot.busy;
            $('pmCatalogSelectedCount').textContent = `已选择 ${state.selectedCount} 条（当前筛选内 ${state.selectedVisible}，筛选外 ${state.selectedHidden}） · 未保存 ${state.draftCount} 条`;
            ['pmCatalogSelectPage', 'pmCatalogSelectFiltered'].forEach(id => { $(id).disabled = unavailable || !snapshot.filtered.length; });
            $('pmCatalogClearSelection').disabled = unavailable || !state.selectedCount;
            $('pmCatalogSaveAll').disabled = unavailable || !state.draftCount || state.results.some(result => result.status === 'unconfirmed');
            root.querySelectorAll('[data-edit-bulk]').forEach(button => { button.disabled = unavailable || !state.selectedAvailable || button.dataset.editBulk === 'clone' && !!state.cloneSubmissionLock; });
            $('pmCatalogCloneLockNotice').classList.toggle('d-none', !state.cloneSubmissionLock);
            $('pmCatalogUnlockClone').disabled = state.busy;
            $('pmCatalogRetryAction').disabled = unavailable || !state.results.some(result => result.status === 'failed');
            $('pmCatalogStopWrites').disabled = !state.canStop;
            $('pmCatalogEditorStopWrites').disabled = !state.canStop;
            $('pmCatalogEditorProgress').textContent = modalBusy ? state.action : '';
            const failed = state.results.filter(result => result.status === 'failed').length, unconfirmed = state.results.filter(result => result.status === 'unconfirmed').length;
            $('pmCatalogActionStatus').textContent = localError || state.configError || `${state.action || (state.ready ? '直接修改所选实际站点；不会经由 Master 修改其他站点。' : '正在读取编辑权限…')}${failed ? ` · 失败 ${failed} 条` : ''}${unconfirmed ? ` · 待核对 ${unconfirmed} 条` : ''}`;
            $('pmCatalogActionStatus').classList.toggle('text-warning', !!(localError || state.configError || failed || unconfirmed));
            $('pmCatalogCloneJobs').innerHTML = state.jobs.map(job => {
                const failures = job.failed_details || [], warnings = job.success_warnings || [];
                const partialNotice = failures.filter(item => item.partial_clone && item.target_id).slice(0, 5).map(item => `<div class="text-warning">已有目标 #${escape(item.target_id)}，变体 / 图片未完整，请核对现有目标，勿重复克隆。</div>`).join('');
                const details = [...failures.map(item => `<div class="text-warning">源商品 #${escape(item.product_id)}：${item.partial_clone && item.target_id ? `已有目标 #${escape(item.target_id)}，变体 / 图片未完整，请核对现有目标，勿重复克隆。` : escape(item.error || '失败')}${item.warnings.length ? ' ' + item.warnings.map(escape).join('；') : ''}</div>`),
                    ...warnings.map(item => `<div>源商品 #${escape(item.product_id)} → 目标 #${escape(item.target_id)}：${item.warnings.map(escape).join('；')}</div>`)].join('');
                return `<div class="small ${job.failed_count || job.poll_error ? 'text-warning' : 'text-muted'}">克隆任务 ${escape(job.job_id)}：${escape(job.status || '等待执行')} · ${Number(job.completed_count || 0)} / ${Number(job.total_count || 0)}，创建 ${Number(job.created_count || 0)}，跳过 ${Number(job.skipped_count || 0)}，失败 ${Number(job.failed_count || 0)}${job.poll_error ? ' · 进度查询失败：' + escape(job.poll_error) + '（任务已提交，请勿重复提交）' : ''}${job.last_error ? ' · ' + escape(job.last_error) : ''}${partialNotice}${details ? '<details class="mt-1"><summary>查看失败 / 警告（最多 50 条失败、20 条警告）</summary>' + details + '</details>' : ''}</div>`;
            }).join('');
        }
        function patchRow(key) {
            const element = [...root.querySelectorAll('[data-catalog-key]')].find(row => row.dataset.catalogKey === key), row = rowByKey(key);
            if (!element || !row) return;
            const values = valuesFor(row), dirty = Object.keys(editor.getDraft(key)).length > 0;
            element.classList.toggle('pm-row-modified', dirty);
            element.querySelector('[data-edit-save]').disabled = !dirty;
            element.querySelector('[data-edit-discard]').disabled = !dirty;
            element.querySelector('[data-edit-feedback]').textContent = resultLabel(key);
            const quantity = element.querySelector('[data-edit-field="stock_quantity"]'), status = element.querySelector('[data-edit-field="stock_status"]');
            quantity.disabled = values.manage_stock !== true;
            status.disabled = values.manage_stock === true || inherited(row);
            if (document.activeElement !== quantity) quantity.value = str(values.stock_quantity);
        }
        async function perform(action) {
            localError = '';
            try { await action(); } catch (error) { localError = error.message || str(error); }
            if (options.onChange) options.onChange();
        }
        function showModal(operation) {
            const state = editor.getSnapshot(), snapshot = catalog.getSnapshot();
            modalKeys = state.selected.slice(); modalOperation = operation;
            const rows = snapshot.rows.filter(row => modalKeys.includes(rowKey(row))), inheritedCount = rows.filter(inherited).length;
            const labels = {price: '批量修改价格', stock: '批量修改库存数量', stock_status: '批量修改库存状态', soldout: '批量硬售完', restore: '恢复数量库存', clear_sale: '清空优惠价', clone: '跨站克隆产品'};
            $('pmCatalogEditorTitle').textContent = labels[operation];
            const parentCount = new Set(rows.map(row => `${row.site_id}:${row.product_id}`)).size;
            let html = `<p class="small text-info">已选 ${rows.length} 条，涉及 ${new Set(rows.map(row => row.site_id)).size} 个实际站点；其中 ${state.selectedHidden} 条在当前筛选之外。${operation === 'clone' ? '' : '仅修改这些已选叶节点，不扩大到同款其他口味。'}</p>`;
            const field = (name, label, attributes, placeholder = '') => `<label class="form-label small mt-2" for="pmCatalogModal-${name}">${label}</label><input id="pmCatalogModal-${name}" name="${name}" class="form-control bg-dark text-white border-secondary" ${attributes} placeholder="${escape(placeholder)}">`;
            if (operation === 'price') html += '<p class="small text-muted">按各站原币种分别设置相同数值，不自动换算；留空保持原值，优惠价填 0 表示清空。</p>' + field('regular_price', '原价', 'type="text" inputmode="decimal" maxlength="64"', '留空保持不变') + field('sale_price', '优惠价', 'type="text" inputmode="decimal" maxlength="64"', '留空保持不变；0 清空');
            else if (operation === 'stock' || operation === 'restore') {
                html += '<p class="small text-muted">启用所选产品 / 口味的数量库存管理，设置为填写的数量。恢复库存不会猜测原来的数量，请填写实际数量。</p>' + field('stock_quantity', '库存数量', 'type="number" min="0" max="2147483647" step="1"', '填写大于等于 0 的整数');
                if (inheritedCount) html += `<div class="alert alert-warning mt-3 small">有 ${inheritedCount} 个口味沿用整款库存。设置后这些口味改为独立库存；整款和未选口味库存不修改。</div><label class="form-check"><input type="checkbox" name="confirm_independent" class="form-check-input">我确认将这些口味改为独立库存，并使用上面填写的数量</label>`;
            } else if (operation === 'stock_status') html += '<p class="small text-muted">关闭所选叶节点的数量管理，使用以下库存状态。</p><label class="form-label" for="pmCatalogModal-stock_status">库存状态</label><select name="stock_status" id="pmCatalogModal-stock_status" class="form-select bg-dark text-white border-secondary">' + Object.entries(STOCK_LABELS).map(([value, label]) => `<option value="${value}">${label}</option>`).join('') + '</select>';
            else if (operation === 'soldout') html += '<p>关闭所选叶节点的数量库存管理，并设为售完。</p>';
            else if (operation === 'clear_sale') html += '<p>将所选产品 / 口味的优惠价清空，恢复使用原价。</p>';
            else if (operation === 'clone') {
                html += `<div class="alert alert-info small">所选口味按整款产品克隆，共 ${parentCount} 款（同站同款去重），每个源站分开创建后台任务。勾选克隆变体会包含整款全部口味。目标站自身的产品不会重复克隆。关闭页面不会中断已提交任务。</div>`;
                html += '<label class="form-label" for="pmCatalogModal-target_site_id">目标站点</label><select name="target_site_id" id="pmCatalogModal-target_site_id" class="form-select bg-dark text-white border-secondary"><option value="">请选择目标站点</option>' + state.config.clone_targets.map(site => `<option value="${Number(site.id)}">${escape((site.manager ? '[' + site.manager + '] ' : '') + site.url)}</option>`).join('') + '</select>';
                html += '<label class="form-check mt-3"><input class="form-check-input" type="checkbox" name="include_variations" checked>克隆全部变体 / 口味</label><label class="form-check"><input class="form-check-input" type="checkbox" name="include_images" checked>克隆图片</label><label class="form-label mt-2" for="pmCatalogModal-status_on_target">目标商品状态</label><select id="pmCatalogModal-status_on_target" name="status_on_target" class="form-select bg-dark text-white border-secondary"><option value="draft">草稿</option><option value="pending">待审核</option><option value="private">私密</option><option value="publish">发布</option></select><label class="form-label mt-2" for="pmCatalogModal-collision_mode">相同 SKU 处理</label><select name="collision_mode" id="pmCatalogModal-collision_mode" class="form-select bg-dark text-white border-secondary"><option value="skip_existing">跳过已存在商品</option><option value="clone_as_new">作为全新商品克隆（新 SKU，强制草稿）</option></select>';
            }
            const blocked = inheritedCount && ['soldout', 'stock_status'].includes(operation);
            if (blocked) html += `<div class="alert alert-warning small mt-3">所选包含 ${inheritedCount} 个沿用整款库存的口味，不能独立执行此动作。请先取消选择这些行，或通过“批量库存数量”明确转为独立库存。</div>`;
            $('pmCatalogEditorBody').innerHTML = html;
            $('pmCatalogEditorError').textContent = '';
            $('pmCatalogEditorApply').disabled = !!blocked;
            $('pmCatalogEditorApply').textContent = operation === 'clone' ? '创建克隆任务' : '确认修改所选结果';
            global.bootstrap.Modal.getOrCreateInstance($('pmCatalogEditorModal')).show();
        }
        root.addEventListener('change', event => {
            const input = event.target, tr = input.closest('[data-catalog-key]');
            if (input.matches('[data-edit-select]')) { editor.select([input.dataset.editSelect], input.checked); return; }
            if (!tr || !input.dataset.editField || !['manage_stock', 'stock_status'].includes(input.dataset.editField)) return;
            const row = rowByKey(tr.dataset.catalogKey);
            if (input.dataset.editField === 'manage_stock' && input.checked && inherited(row) && !global.confirm('此口味目前沿用整款库存。启用后改为独立库存，需要另行填写该口味数量；整款和未选口味不会修改。是否继续？')) { input.checked = false; return; }
            try { editor.setDraft(tr.dataset.catalogKey, input.dataset.editField, input.type === 'checkbox' ? input.checked : input.value); patchRow(tr.dataset.catalogKey); }
            catch (error) { localError = error.message; if (options.onChange) options.onChange(); }
        });
        root.addEventListener('input', event => {
            const input = event.target, tr = input.closest('[data-catalog-key]');
            if (!tr || !['stock_quantity', 'regular_price', 'sale_price'].includes(input.dataset.editField)) return;
            try { editor.setDraft(tr.dataset.catalogKey, input.dataset.editField, input.value); patchRow(tr.dataset.catalogKey); }
            catch (error) { localError = error.message; if (options.onChange) options.onChange(); }
        });
        root.addEventListener('click', event => {
            const button = event.target.closest('button'); if (!button || button.disabled) return;
            if (button.dataset.editSave) perform(() => editor.saveDrafts([button.dataset.editSave]));
            else if (button.dataset.editDiscard) editor.discard(button.dataset.editDiscard);
            else if (button.dataset.editVerify) perform(() => editor.verify(button.dataset.editVerify));
            else if (button.dataset.editBulk) showModal(button.dataset.editBulk);
        });
        $('pmCatalogSelectPage').addEventListener('click', () => editor.select(catalog.getSnapshot().pageRows.map(rowKey)));
        $('pmCatalogSelectFiltered').addEventListener('click', () => editor.select(catalog.getSnapshot().filtered.map(rowKey)));
        $('pmCatalogClearSelection').addEventListener('click', editor.clearSelection);
        $('pmCatalogSaveAll').addEventListener('click', () => perform(() => editor.saveDrafts()));
        $('pmCatalogRetryAction').addEventListener('click', () => perform(editor.retryFailed));
        $('pmCatalogStopWrites').addEventListener('click', editor.stopWrites);
        $('pmCatalogEditorStopWrites').addEventListener('click', editor.stopWrites);
        $('pmCatalogEditorBody').addEventListener('change', event => {
            if (event.target.name === 'collision_mode') {
                const status = $('pmCatalogModal-status_on_target'); status.disabled = event.target.value === 'clone_as_new'; if (status.disabled) status.value = 'draft';
            }
        });
        $('pmCatalogEditorApply').addEventListener('click', async () => {
            if (modalBusy) return;
            const values = {};
            $('pmCatalogEditorBody').querySelectorAll('[name]').forEach(input => { values[input.name] = input.type === 'checkbox' ? input.checked : input.value; });
            if (!sameIdentity([...editor.getSnapshot().selected].sort(), [...modalKeys].sort())) { $('pmCatalogEditorError').textContent = '选择范围已变化，请关闭弹框后重新确认。'; return; }
            modalBusy = true; $('pmCatalogEditorApply').disabled = true; $('pmCatalogEditorError').textContent = '';
            $('pmCatalogEditorModal').querySelectorAll('[data-bs-dismiss]').forEach(button => { button.disabled = true; });
            try {
                if (modalOperation === 'clone') await editor.cloneSelected(values); else await editor.bulk(modalOperation, values);
                global.bootstrap.Modal.getInstance($('pmCatalogEditorModal')).hide();
            } catch (error) { $('pmCatalogEditorError').textContent = error.message || str(error); }
            finally {
                modalBusy = false; $('pmCatalogEditorApply').disabled = modalOperation === 'clone' && !!editor.getSnapshot().cloneSubmissionLock;
                $('pmCatalogEditorModal').querySelectorAll('[data-bs-dismiss]').forEach(button => { button.disabled = false; });
            }
        });
        $('pmCatalogUnlockClone').addEventListener('click', editor.acknowledgeCloneSubmission);
        return {editor, render, tableHtml, get tableVersion() { return editor.getSnapshot().tableVersion; }, canReload: editor.canReload, getSnapshot: editor.getSnapshot};
    }

    const api = {createEditor, createUI, validChanges, sameField, sameIdentity, rowKey, inherited, escape};
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    global.ProductManagerCatalogEditor = api;
})(typeof window === 'undefined' ? globalThis : window);
