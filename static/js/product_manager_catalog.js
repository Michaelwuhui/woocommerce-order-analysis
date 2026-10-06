(function (global) {
    'use strict';

    const PAGE_SIZE = 100;
    const FILTER_IDS = {
        manager: 'pmCatalogManagerFilter', site: 'pmCatalogSiteFilter',
        brand: 'pmCatalogBrandFilter', product: 'pmCatalogProductFilter',
        flavor: 'pmCatalogFlavorFilter', stock: 'pmCatalogStockFilter',
        publish: 'pmCatalogPublishFilter', text: 'pmCatalogTextFilter'
    };
    const STOCK_LABELS = {instock: '有货', outofstock: '售完', onbackorder: '可预订', unknown: '未知'};
    const PUBLISH_LABELS = {publish: '已发布', draft: '草稿', private: '私密', pending: '待审核', future: '定时发布', trash: '回收站', 'auto-draft': '自动草稿', unknown: '未知'};
    const SITE_LABELS = {pending: '等待读取', loading: '读取中', complete: '完整', failed: '失败 / 不完整', incomplete: '不完整', stopped: '已停止 / 不完整'};
    const str = value => value === undefined || value === null ? '' : String(value);
    const normalized = value => str(value).normalize('NFKC').trim().toLocaleLowerCase();
    const searchNormalized = value => str(value).replace(/<[^>]*>/g, ' ').toLocaleLowerCase()
        .replace(/[łøđðþæœß]/g, character => ({ł: 'l', ø: 'o', đ: 'd', ð: 'd', þ: 'th', æ: 'ae', œ: 'oe', ß: 'ss'}[character]))
        .normalize('NFKD').replace(/\p{M}/gu, '').replace(/[^\p{L}\p{N}]+/gu, ' ').replace(/\s+/g, ' ').trim();
    const escapeHtml = value => str(value).replace(/[&<>"']/g, character => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[character]));
    const list = value => Array.isArray(value) ? [...new Set(value.map(item => str(item && typeof item === 'object' ? item.name || item.slug : item).trim()).filter(Boolean))] : [];
    const valueKey = value => 'value:' + normalized(value);
    const stockValue = row => str(row.stock_state || row.stock_status || 'unknown');
    const managerLabel = site => str(site.manager).trim() || '未分配负责人';
    const siteLabel = site => str(site.name || site.url || ('站点 ' + site.id));
    const emptyFilters = () => ({manager: '', site: '', brand: '', product: '', flavor: '', stock: '', publish: '', text: ''});

    function safeUrl(value) {
        try {
            const parsed = new URL(str(value));
            return parsed.protocol === 'http:' || parsed.protocol === 'https:' ? parsed.href : '';
        } catch (_) { return ''; }
    }

    function normalizeRow(row, site) {
        const productId = Number(row.product_id), variationId = Number(row.variation_id || 0);
        if (!Number.isSafeInteger(productId) || productId <= 0 || !Number.isSafeInteger(variationId) || variationId < 0) {
            throw new Error('接口返回无效的产品或变体编号，当前站点结果不完整。');
        }
        if (row.site_id !== undefined && str(row.site_id) !== str(site.id)) throw new Error('产品所属站点不符，当前站点结果不完整。');
        return Object.assign({}, row, {
            site_id: str(site.id), product_id: productId, variation_id: variationId,
            key: `${site.id}:${productId}:${variationId}`, site: site,
            product_name: str(row.product_name || row.name), name: str(row.name),
            sku: str(row.sku), brands: list(row.brands), flavors: list(row.flavors)
        });
    }

    function matchesFilters(row, filters) {
        return (!filters.manager || valueKey(managerLabel(row.site)) === filters.manager) &&
            (!filters.site || str(row.site_id) === filters.site) &&
            (!filters.brand || (row.brands.length ? row.brands.some(value => valueKey(value) === filters.brand) : filters.brand === 'missing:brand')) &&
            (!filters.product || valueKey(row.product_name) === filters.product) &&
            (!filters.flavor || (row.flavors.length ? row.flavors.some(value => valueKey(value) === filters.flavor) : filters.flavor === 'missing:flavor')) &&
            (!filters.stock || stockValue(row) === filters.stock) &&
            (!filters.publish || str(row.status || 'unknown') === filters.publish) &&
            (!searchNormalized(filters.text) || searchNormalized([siteLabel(row.site), managerLabel(row.site), row.product_name, row.name, row.sku, ...row.brands, ...row.flavors,
                ...(Array.isArray(row.attributes) ? row.attributes.flatMap(attribute => list(attribute.values)) : [])].join(' ')).includes(searchNormalized(filters.text)));
    }

    function facetOptions(rows, sites) {
        const maps = Object.fromEntries(['manager', 'site', 'brand', 'product', 'flavor', 'stock', 'publish'].map(key => [key, new Map()]));
        sites.forEach(site => {
            maps.manager.set(valueKey(managerLabel(site)), managerLabel(site));
            maps.site.set(str(site.id), siteLabel(site));
        });
        rows.forEach(row => {
            row.brands.forEach(value => maps.brand.set(valueKey(value), value));
            if (!row.brands.length) maps.brand.set('missing:brand', '未识别品牌');
            maps.product.set(valueKey(row.product_name), row.product_name || '未提供产品名');
            row.flavors.forEach(value => maps.flavor.set(valueKey(value), value));
            if (!row.flavors.length) maps.flavor.set('missing:flavor', '未识别口味');
            const stock = stockValue(row), publish = str(row.status || 'unknown');
            maps.stock.set(stock, STOCK_LABELS[stock] || stock);
            maps.publish.set(publish, PUBLISH_LABELS[publish] || publish);
        });
        return Object.fromEntries(Object.entries(maps).map(([key, map]) => [key, [...map.entries()].sort((a, b) => a[1].localeCompare(b[1], 'zh-CN')).map(([value, label]) => ({value, label}))]));
    }

    async function readJson(fetcher, url, signal, timeoutMs) {
        const requestAborter = new AbortController();
        const cancel = () => requestAborter.abort();
        let timedOut = false;
        signal.addEventListener('abort', cancel, {once: true});
        if (signal.aborted) cancel();
        // A variation page reads its parent and one WC page; the browser deadline
        // allows both bounded upstream requests before ending a stuck Web request.
        const timer = setTimeout(() => { timedOut = true; requestAborter.abort(); }, timeoutMs || 65000);
        try {
            let response;
            try { response = await fetcher(url, {method: 'GET', signal: requestAborter.signal, headers: {Accept: 'application/json'}}); }
            catch (error) {
                if (signal.aborted) throw error;
                if (timedOut) throw new Error('当前分页读取超时，该站点结果不完整，可稍后单独重试。');
                throw new Error('网络请求失败：' + (error.message || '无法连接服务器'));
            }
            const contentType = response.headers && response.headers.get('content-type') || '';
            if (!/\bapplication\/(?:[\w.+-]+\+)?json\b/i.test(contentType)) {
                throw new Error(`接口返回了非 JSON 内容（HTTP ${response.status}），可能是登录过期、代理错误或站点异常。`);
            }
            let data;
            try { data = await response.json(); }
            catch (_) {
                if (timedOut) throw new Error('当前分页读取超时，该站点结果不完整，可稍后单独重试。');
                throw new Error(`接口 JSON 无法解析（HTTP ${response.status}）。`);
            }
            if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('接口 JSON 格式无效，无法确认读取结果。');
            if (!response.ok || data.error || data.success === false) {
                throw new Error(`HTTP ${response.status}：${str(data.error || data.message || '接口请求失败')}`);
            }
            return data;
        } finally {
            clearTimeout(timer);
            signal.removeEventListener('abort', cancel);
        }
    }

    function createController(options) {
        const fetcher = options.fetch;
        const concurrency = Math.max(1, Math.min(3, Number(options.concurrency) || 3));
        let epoch = 0, aborter = null;
        const state = {phase: 'idle', queryMode: 'flavor', brand: '', keyword: '', sites: [], siteStates: new Map(), rows: new Map(), filters: emptyFilters(), page: 1, error: ''};
        function snapshot() {
            const rows = [...state.rows.values()];
            const filtered = rows.filter(row => matchesFilters(row, state.filters));
            const siteStates = state.sites.map(site => Object.assign({}, state.siteStates.get(str(site.id))));
            const complete = siteStates.filter(site => site.status === 'complete').length;
            const failed = siteStates.filter(site => site.status === 'failed').length;
            const incomplete = siteStates.filter(site => ['failed', 'incomplete', 'stopped'].includes(site.status)).length;
            const pages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
            state.page = Math.max(1, Math.min(state.page, pages));
            return {phase: state.phase, queryMode: state.queryMode, mode: state.queryMode, brand: state.brand, keyword: state.keyword, sites: state.sites.slice(), siteStates,
                rows, filtered, filters: Object.assign({}, state.filters), options: facetOptions(rows, state.sites),
                page: state.page, pages, pageRows: filtered.slice((state.page - 1) * PAGE_SIZE, state.page * PAGE_SIZE),
                counts: {matched: rows.length, filtered: filtered.length, totalSites: state.sites.length, completeSites: complete, failedSites: failed, incompleteSites: incomplete}, error: state.error,
                busy: state.phase === 'discovering' || state.phase === 'loading'};
        }
        function emit() { if (options.onChange) options.onChange(snapshot()); }
        function current(run) { return run === epoch && aborter && !aborter.signal.aborted; }
        function begin() {
            if (aborter) aborter.abort();
            aborter = new AbortController();
            epoch += 1;
            return epoch;
        }
        function stop() {
            if (!['discovering', 'loading'].includes(state.phase)) return;
            if (aborter) aborter.abort();
            epoch += 1;
            state.siteStates.forEach(site => { if (site.status === 'loading' || site.status === 'pending') site.status = 'stopped'; });
            state.phase = 'stopped';
            emit();
        }
        async function scanPages(siteState, parentId, run, parents) {
            let page = 1;
            let knownTotal = null, knownPages = null;
            const sourceIds = new Set(), sourceStatuses = new Map(), requestedPages = new Set();
            function knownNumber(value, label) {
                if (value === null || value === undefined) return null;
                if (!Number.isSafeInteger(Number(value)) || Number(value) < 0) throw new Error(label + '无效，当前站点结果不完整。');
                return Number(value);
            }
            while (current(run)) {
                if (requestedPages.has(page)) throw new Error('分页重复，当前站点结果不完整。');
                requestedPages.add(page);
                const params = new URLSearchParams({site_id: siteState.site.id, parent_id: parentId, page, search: state.keyword, brand: state.brand, query_mode: state.queryMode});
                const data = await readJson(fetcher, '/api/product-manager/catalog-page?' + params.toString(), aborter.signal, options.requestTimeout);
                if (!current(run)) return;
                if (data.complete_page !== true || !Array.isArray(data.rows) || !Array.isArray(data.source_ids) || typeof data.has_more !== 'boolean') {
                    throw new Error('分页结果缺少完整性信息，当前站点结果不完整。');
                }
                if (Number(data.page) !== page || Number(data.parent_id) !== parentId || str(data.site_id) !== str(siteState.site.id)) {
                    throw new Error('接口返回了其他站点或分页的数据，当前站点结果不完整。');
                }
                const total = knownNumber(data.total, '分页总数'), pages = knownNumber(data.total_pages, '总页数');
                if ((knownTotal !== null && total !== null && knownTotal !== total) ||
                    (knownPages !== null && pages !== null && knownPages !== pages)) {
                    throw new Error('读取期间分页总数发生变化，当前站点结果不完整，请刷新后重试。');
                }
                if (total !== null) knownTotal = total;
                if (pages !== null) knownPages = pages;
                if (Number(data.scanned) !== data.source_ids.length) throw new Error('分页读取数量与来源编号不符，当前站点结果不完整。');
                const pageIds = new Set();
                data.source_ids.forEach(id => {
                    const numericId = Number(id), key = str(numericId);
                    if (!Number.isSafeInteger(numericId) || numericId <= 0) throw new Error('分页来源编号无效，当前站点结果不完整。');
                    if (sourceIds.has(key)) throw new Error('站点分页重复返回同一产品或变体，当前站点结果不完整，请刷新后重试。');
                    sourceIds.add(key);
                    pageIds.add(key);
                });
                if (parentId) {
                    const statuses = data.source_statuses;
                    if (!statuses || typeof statuses !== 'object' || Array.isArray(statuses) || Object.keys(statuses).length !== pageIds.size) {
                        throw new Error('变体分页缺少完整发布状态，当前站点结果不完整。');
                    }
                    pageIds.forEach(id => {
                        const status = statuses[id];
                        if (!Object.prototype.hasOwnProperty.call(statuses, id) || typeof status !== 'string' || !Object.prototype.hasOwnProperty.call(PUBLISH_LABELS, status) || status === 'unknown') {
                            throw new Error('变体分页含缺失或未知发布状态，当前站点结果不完整。');
                        }
                        sourceStatuses.set(id, status);
                    });
                }
                // Normalize the entire page first: malformed pages never partly masquerade as valid pages.
                const incoming = data.rows.map(row => normalizeRow(row, siteState.site));
                incoming.forEach(row => {
                    if (!pageIds.has(str(parentId ? row.variation_id : row.product_id)) ||
                        (parentId ? row.product_id !== parentId || row.variation_id === 0 : row.variation_id !== 0)) {
                        throw new Error('产品或变体与当前分页不符，当前站点结果不完整。');
                    }
                });
                if (!parentId) {
                    if (!Array.isArray(data.variable_products)) throw new Error('缺少可变产品列表，当前站点结果不完整。');
                    data.variable_products.forEach(parent => {
                        const id = Number(parent.id);
                        if (!Number.isSafeInteger(id) || id <= 0) throw new Error('可变产品编号无效，当前站点结果不完整。');
                        parents.set(id, parent);
                    });
                    siteState.productPages += 1;
                    siteState.productTotalPages = data.total_pages;
                } else {
                    siteState.variationPages += 1;
                }
                incoming.forEach(row => state.rows.set(row.key, row));
                siteState.scanned += Number(data.scanned) || 0;
                list(data.warnings).forEach(warning => { if (!siteState.warnings.includes(warning)) siteState.warnings.push(warning); });
                siteState.matched = [...state.rows.values()].filter(row => row.site_id === str(siteState.site.id)).length;
                emit();
                if (!data.has_more) {
                    if ((knownTotal !== null && sourceIds.size !== knownTotal) ||
                        (knownPages !== null && requestedPages.size !== Math.max(1, knownPages))) {
                        throw new Error('已读取数量与站点声明的目录总数不一致，当前站点结果不完整，请刷新后重试。');
                    }
                    return {sourceIds, sourceStatuses};
                }
                const nextPage = Number(data.next_page);
                if (!Number.isSafeInteger(nextPage) || nextPage !== page + 1 || data.source_ids.length === 0) {
                    throw new Error('分页没有向前推进，当前站点结果不完整。');
                }
                page = nextPage;
            }
        }
        async function scanSite(siteState, run) {
            if (!current(run)) return;
            siteState.status = 'loading';
            emit();
            try {
                const parents = new Map();
                await scanPages(siteState, 0, run, parents);
                siteState.variableTotal = parents.size;
                for (const parentId of parents.keys()) {
                    if (!current(run)) return;
                    siteState.currentProduct = str(parents.get(parentId).name || ('产品 ' + parentId));
                    const scanned = await scanPages(siteState, parentId, run, parents);
                    if (!current(run)) return;
                    const expectedIds = parents.get(parentId).variation_ids;
                    if (!Array.isArray(expectedIds)) throw new Error('可变产品缺少完整变体编号，当前站点结果不完整。');
                    const expected = new Set(expectedIds.map(id => str(Number(id))));
                    // Woo's parent get_children() lists publish/private children.
                    // status=any also returns drafts and scheduled variations;
                    // those remain valid rows without expanding the parent list.
                    const seen = new Set([...scanned.sourceStatuses.entries()].filter(([, status]) => status === 'publish' || status === 'private').map(([id]) => id));
                    if (expected.size !== expectedIds.length || [...expected].some(id => !Number.isSafeInteger(Number(id)) || Number(id) <= 0) ||
                        seen.size !== expected.size || [...expected].some(id => !seen.has(id))) {
                        throw new Error('读取期间产品变体目录发生变化或缺失，当前站点结果不完整，请刷新后重试。');
                    }
                    siteState.variableDone += 1;
                    emit();
                }
                if (!current(run)) return;
                siteState.currentProduct = '';
                siteState.status = siteState.warnings.length ? 'incomplete' : 'complete';
            } catch (error) {
                if (!current(run)) return;
                siteState.status = 'failed';
                siteState.error = error.message || str(error);
            }
            emit();
        }
        function newSiteState(site) {
            return {site, status: 'pending', error: '', warnings: [], productPages: 0, productTotalPages: null,
                variationPages: 0, variableTotal: 0, variableDone: 0, currentProduct: '', scanned: 0, matched: 0};
        }
        async function runSites(targets, run) {
            state.phase = 'loading';
            emit();
            let next = 0;
            async function worker() {
                while (current(run) && next < targets.length) {
                    const site = targets[next++];
                    await scanSite(state.siteStates.get(str(site.id)), run);
                }
            }
            await Promise.all(Array.from({length: Math.min(concurrency, targets.length)}, worker));
            if (!current(run)) return;
            state.phase = [...state.siteStates.values()].every(site => site.status === 'complete') ? 'complete' : 'incomplete';
            emit();
        }
        async function load(keyword, loadOptions) {
            const run = begin();
            state.queryMode = loadOptions && loadOptions.queryMode || 'flavor';
            state.brand = state.queryMode === 'brand' ? str(loadOptions && loadOptions.brand).trim() : '';
            state.keyword = state.queryMode === 'flavor' ? str(keyword).trim() : '';
            state.phase = 'discovering';
            state.error = '';
            state.rows.clear();
            state.sites = [];
            state.siteStates.clear();
            state.page = 1;
            if (!loadOptions || !loadOptions.preserveFilters) state.filters = emptyFilters();
            emit();
            if (!['flavor', 'brand', 'all'].includes(state.queryMode) || (state.queryMode === 'brand' && !state.brand)) {
                state.phase = 'error';
                state.error = state.queryMode === 'brand' ? '请输入或选择一个品牌，再加载全部站点。' : '请选择有效的查询方式。';
                emit();
                return;
            }
            if (state.brand.length > 200 || state.keyword.length > 200) {
                state.phase = 'error';
                state.error = '品牌或口味关键词不能超过 200 个字符。';
                emit();
                return;
            }
            try {
                const data = await readJson(fetcher, '/api/product-manager/catalog-sites', aborter.signal, options.requestTimeout);
                if (!current(run)) return;
                if (!Array.isArray(data.sites)) throw new Error('站点列表格式无效，无法确认查询范围。');
                const ids = new Set();
                state.sites = data.sites.map(site => {
                    const id = Number(site.id);
                    if (!Number.isSafeInteger(id) || id <= 0 || ids.has(id)) throw new Error('站点列表有无效或重复编号，无法确认查询范围。');
                    ids.add(id);
                    return Object.assign({}, site, {id: str(id)});
                });
                state.sites.forEach(site => state.siteStates.set(str(site.id), newSiteState(site)));
                await runSites(state.sites, run);
            } catch (error) {
                if (!current(run)) return;
                state.phase = 'error';
                state.error = error.message || str(error);
                emit();
            }
        }
        async function retryIncomplete() {
            if (['discovering', 'loading'].includes(state.phase)) return;
            const targets = state.sites.filter(site => state.siteStates.get(str(site.id)).status !== 'complete');
            if (!targets.length) return;
            const run = begin();
            state.error = '';
            const targetIds = new Set(targets.map(site => str(site.id)));
            state.rows.forEach((row, key) => { if (targetIds.has(row.site_id)) state.rows.delete(key); });
            targets.forEach(site => state.siteStates.set(str(site.id), newSiteState(site)));
            await runSites(targets, run);
        }
        return {load, stop, retryIncomplete, getSnapshot: snapshot,
            loadQuery(query) { return load(query.keyword || query.search || '', {queryMode: query.queryMode || query.mode || 'flavor', brand: query.brand, preserveFilters: !!query.preserveFilters}); },
            refresh() { return load(state.keyword, {queryMode: state.queryMode, brand: state.brand, preserveFilters: true}); },
            setFilters(filters) { state.filters = Object.assign({}, state.filters, filters); state.page = 1; emit(); },
            resetFilters() { state.filters = emptyFilters(); state.page = 1; emit(); },
            setPage(page) { state.page = Number(page) || 1; emit(); }};
    }

    function resultHtml(snapshot) {
        if (!snapshot.pageRows.length) {
            let message;
            if (snapshot.phase === 'idle') message = '选择按品牌、按口味或全部产品，点击「加载全部站点」开始查询。';
            else if (snapshot.counts.matched) message = '已加载的匹配结果中没有符合当前筛选条件的产品。可清空筛选后查看。';
            else if (snapshot.busy) message = '正在逐站读取，暂未收到匹配结果，请查看读取进度。';
            else if (snapshot.phase === 'complete' && snapshot.counts.totalSites) message = '已完整读取全部有权限站点，此次查询未匹配到产品或口味。';
            else if (snapshot.phase === 'complete') message = '当前账号没有可查询的站点，请联系管理员核对产品管理权限。';
            else message = '当前结果尚不完整，暂未匹配到产品。不能据此判断全部站点没有符合查询条件的产品，请重试异常站点或重新加载。';
            return '<div class="pm-empty-state">' + escapeHtml(message) + '</div>';
        }
        const rows = snapshot.pageRows.map(row => {
            const url = safeUrl(row.permalink), stock = stockValue(row);
            const quantity = row.manage_stock ? (row.stock_quantity === null || row.stock_quantity === undefined ? '数量未知' : str(row.stock_quantity)) : '未管理数量';
            const currency = str(row.currency || row.site.currency).trim();
            const price = row.price === null || row.price === undefined || row.price === '' ? '未提供价格' : str(row.price) + (currency ? ' ' + currency : '');
            const stockClass = stock === 'outofstock' ? 'text-danger' : stock === 'onbackorder' ? 'text-warning' : '';
            return `<tr data-catalog-key="${escapeHtml(row.key)}"><td>${escapeHtml(siteLabel(row.site))}<div class="text-muted small">${escapeHtml(managerLabel(row.site))}</div></td>` +
                `<td>${escapeHtml(row.brands.join(' / ') || '未识别')}</td><td class="pm-name">${url ? `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(row.product_name)}</a>` : escapeHtml(row.product_name)}<div class="text-muted small">${escapeHtml(PUBLISH_LABELS[row.status] || row.status || '未知')} · ${row.variation_id ? '变体 ' + row.variation_id : '产品 ' + row.product_id}</div></td>` +
                `<td>${escapeHtml(row.flavors.join(' / ') || '未识别口味')}${row.flavor_scope === 'any' ? '<div class="small text-info">任意口味（共用此变体）</div>' : ''}${row.variation_id && row.name !== row.product_name ? '<div class="text-muted small">' + escapeHtml(row.name) + '</div>' : ''}<div class="pm-sku">SKU：${escapeHtml(row.sku || '未提供')}</div></td>` +
                `<td class="${stockClass}">${escapeHtml(STOCK_LABELS[stock] || stock)}<div class="small">${escapeHtml(quantity)}</div></td><td>${escapeHtml(price)}</td>` +
                `<td><button type="button" class="btn btn-sm btn-outline-light" data-catalog-open="${escapeHtml(row.key)}">单站管理</button></td></tr>`;
        }).join('');
        const pagination = `<div class="d-flex justify-content-between align-items-center p-3 gap-2"><span class="text-muted small">第 ${snapshot.page} / ${snapshot.pages} 页，每页最多 ${PAGE_SIZE} 条</span><div class="d-flex gap-2"><button type="button" class="btn btn-sm btn-outline-light" data-catalog-page="${snapshot.page - 1}" ${snapshot.page === 1 ? 'disabled' : ''}>上一页</button><button type="button" class="btn btn-sm btn-outline-light" data-catalog-page="${snapshot.page + 1}" ${snapshot.page === snapshot.pages ? 'disabled' : ''}>下一页</button></div></div>`;
        return '<div class="table-responsive"><table class="table table-dark table-hover pm-table mb-0"><thead><tr><th>站点 / 负责人</th><th>品牌</th><th>产品 / 发布状态</th><th>口味 / SKU</th><th>库存</th><th>价格（站点原币）</th><th>管理</th></tr></thead><tbody>' + rows + '</tbody></table></div>' + pagination;
    }

    function mount(document, mountOptions) {
        const root = document.getElementById('pmCatalogPane');
        if (!root) return null;
        const $ = id => document.getElementById(id);
        let latest, scheduled = false;
        const controller = createController({fetch: mountOptions && mountOptions.fetch || global.fetch.bind(global), onChange(snapshot) {
            latest = snapshot;
            if (scheduled) return;
            scheduled = true;
            (global.requestAnimationFrame || (callback => setTimeout(callback, 0)))(() => { scheduled = false; render(latest); });
        }});
        function render(snapshot) {
            $('pmCatalogLoad').disabled = false; // A new keyword can replace an in-flight query.
            $('pmCatalogStop').disabled = !snapshot.busy;
            $('pmCatalogRetry').disabled = snapshot.busy || !snapshot.counts.incompleteSites;
            $('pmCatalogRefresh').disabled = snapshot.phase === 'idle';
            let status = {idle: '选择查询方式，一次查询全部有权限站点；加载后可继续组合筛选。', discovering: '正在确认当前账号有权限的全部站点…', loading: '正在读取各站产品与全部变体，匹配结果会陆续显示。', complete: '查询完成，全部站点读取完整。', incomplete: '查询结束，存在不完整站点；已保留成功读取的匹配结果。', stopped: '已停止；已读取的匹配结果保留，未完成站点可单独重试。', error: snapshot.error}[snapshot.phase];
            if (snapshot.phase === 'complete' && !snapshot.counts.totalSites) status = '当前账号没有可查询的站点，请核对产品管理权限。';
            if (snapshot.phase === 'stopped' && !snapshot.counts.totalSites) status = '已停止确认站点范围，请重新加载以继续查询。';
            if (snapshot.phase !== 'idle') {
                const scope = snapshot.queryMode === 'brand' ? (snapshot.brand ? `品牌「${snapshot.brand}」` : '品牌查询') : snapshot.queryMode === 'all' || !snapshot.keyword ? '全部产品' : `口味 / 关键词「${snapshot.keyword}」`;
                status = scope + '：' + status;
            }
            $('pmCatalogStatus').textContent = status;
            $('pmCatalogStatus').classList.toggle('text-warning', ['incomplete', 'stopped', 'error'].includes(snapshot.phase));
            const counts = snapshot.counts;
            $('pmCatalogCounts').textContent = `已匹配 ${counts.matched} 条 · 筛选后 ${counts.filtered} 条 · 完整站点 ${counts.completeSites} / ${counts.totalSites} · 失败 ${counts.failedSites} · 不完整 ${counts.incompleteSites}`;
            $('pmCatalogProgress').innerHTML = snapshot.siteStates.map(site => {
                const problems = [site.error, ...site.warnings].filter(Boolean).map(escapeHtml).join('；');
                const progress = `产品 ${site.productPages}${Number.isFinite(Number(site.productTotalPages)) && site.productTotalPages !== null ? ' / ' + Number(site.productTotalPages) : ''} 页 · 已检查 ${site.scanned} 项 · 匹配 ${site.matched} 条 · 变体 ${site.variableDone} / ${site.variableTotal} 个产品（${site.variationPages} 页）`;
                return `<div class="border-bottom border-secondary border-opacity-25 py-2"><div class="d-flex justify-content-between gap-2"><strong>${escapeHtml(siteLabel(site.site))}</strong><span class="${site.status === 'complete' ? 'text-success' : ['failed', 'incomplete', 'stopped'].includes(site.status) ? 'text-warning' : 'text-info'}">${escapeHtml(SITE_LABELS[site.status])}</span></div><div class="small text-muted">${escapeHtml(managerLabel(site.site))} · ${escapeHtml(progress)}${site.currentProduct ? ' · 正在读取：' + escapeHtml(site.currentProduct) : ''}</div>${problems ? '<div class="small text-warning">' + problems + '</div>' : ''}</div>`;
            }).join('');
            Object.entries(FILTER_IDS).forEach(([key, id]) => {
                const element = $(id);
                if (!element) return;
                if (key === 'text') { if (element.value !== snapshot.filters.text) element.value = snapshot.filters.text; return; }
                const selected = snapshot.filters[key];
                const choices = snapshot.options[key].slice();
                if (selected && !choices.some(choice => choice.value === selected)) {
                    const old = [...element.options].find(option => option.value === selected);
                    choices.push({value: selected, label: old ? old.textContent : selected});
                }
                element.innerHTML = '<option value="">全部</option>' + choices.map(choice => `<option value="${escapeHtml(choice.value)}" ${choice.value === selected ? 'selected' : ''}>${escapeHtml(choice.label)}</option>`).join('');
            });
            $('pmCatalogResults').innerHTML = resultHtml(snapshot);
        }
        function syncQueryForm() {
            const mode = $('pmCatalogQueryMode') ? $('pmCatalogQueryMode').value : 'flavor';
            if ($('pmCatalogFlavorInputWrap')) $('pmCatalogFlavorInputWrap').classList.toggle('d-none', mode !== 'flavor');
            if ($('pmCatalogBrandInputWrap')) $('pmCatalogBrandInputWrap').classList.toggle('d-none', mode !== 'brand');
        }
        function loadFormQuery() {
            const queryMode = $('pmCatalogQueryMode') ? $('pmCatalogQueryMode').value : 'flavor';
            return controller.load($('pmCatalogSearch').value, {queryMode, brand: $('pmCatalogBrandInput') ? $('pmCatalogBrandInput').value : ''});
        }
        function restoreLoadedQueryForm() {
            const snapshot = controller.getSnapshot();
            if ($('pmCatalogQueryMode')) $('pmCatalogQueryMode').value = snapshot.queryMode;
            if ($('pmCatalogBrandInput')) $('pmCatalogBrandInput').value = snapshot.brand;
            $('pmCatalogSearch').value = snapshot.keyword;
            syncQueryForm();
        }
        $('pmCatalogLoad').addEventListener('click', loadFormQuery);
        [$('pmCatalogSearch'), $('pmCatalogBrandInput')].filter(Boolean).forEach(input => input.addEventListener('keydown', event => {
            if (event.key === 'Enter') { event.preventDefault(); loadFormQuery(); }
        }));
        if ($('pmCatalogQueryMode')) $('pmCatalogQueryMode').addEventListener('change', syncQueryForm);
        $('pmCatalogRefresh').addEventListener('click', () => { restoreLoadedQueryForm(); controller.refresh(); });
        $('pmCatalogStop').addEventListener('click', controller.stop);
        $('pmCatalogRetry').addEventListener('click', () => { restoreLoadedQueryForm(); controller.retryIncomplete(); });
        $('pmCatalogResetFilters').addEventListener('click', controller.resetFilters);
        Object.entries(FILTER_IDS).forEach(([key, id]) => { if ($(id)) $(id).addEventListener(key === 'text' ? 'input' : 'change', event => controller.setFilters({[key]: event.target.value})); });
        root.addEventListener('click', event => {
            const pageButton = event.target.closest('[data-catalog-page]');
            if (pageButton && !pageButton.disabled) { controller.setPage(pageButton.dataset.catalogPage); return; }
            const openButton = event.target.closest('[data-catalog-open]');
            if (!openButton) return;
            const row = controller.getSnapshot().rows.find(item => item.key === openButton.dataset.catalogOpen);
            if (row) root.dispatchEvent(new CustomEvent('pm:open-single-site', {bubbles: true, detail: {site_id: row.site_id, search: row.product_name}}));
        });
        render(controller.getSnapshot());
        syncQueryForm(); // Preserve the template's initial mode, including its brand default.
        root.productManagerCatalog = controller;
        return controller;
    }

    const api = {createController, mount, escapeHtml, safeUrl, matchesFilters, facetOptions, resultHtml};
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    if (global.document) {
        global.ProductManagerCatalog = api;
        if (global.document.readyState === 'loading') global.document.addEventListener('DOMContentLoaded', () => mount(global.document));
        else mount(global.document);
    }
})(typeof window === 'undefined' ? globalThis : window);
