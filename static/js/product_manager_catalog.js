(function (global) {
    'use strict';

    const PAGE_SIZE = 100;
    const FILTER_IDS = {
        manager: 'pmCatalogManagerFilter', country: 'pmCatalogCountryFilter', site: 'pmCatalogSiteFilter',
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
    const countryCode = row => str(row.country === undefined ? row.site && row.site.country : row.country).trim();
    const countryKey = value => str(value).trim() ? valueKey(value) : 'missing:country';
    const emptyFilters = () => ({manager: '', country: '', site: '', brand: '', product: '', flavor: '', stock: '', publish: '', text: ''});
    const emptyScope = () => ({manager: '', managerMode: 'all', country: '', countryMode: 'all', siteId: ''});
    function normalizeScope(input, previous) {
        const scope = Object.assign(emptyScope(), previous || {}, input || {});
        ['manager', 'country'].forEach(key => {
            scope[key] = str(scope[key]).trim();
            if (input && Object.prototype.hasOwnProperty.call(input, key) && !Object.prototype.hasOwnProperty.call(input, key + 'Mode')) scope[key + 'Mode'] = scope[key] ? 'exact' : 'all';
            if (!['all', 'exact', 'unassigned'].includes(scope[key + 'Mode'])) throw new Error('加载范围方式无效。');
            if (scope[key + 'Mode'] !== 'exact') scope[key] = '';
        });
        scope.siteId = str(scope.siteId).trim();
        if (scope.siteId && (!Number.isSafeInteger(Number(scope.siteId)) || Number(scope.siteId) <= 0)) throw new Error('加载站点编号无效。');
        if (scope.siteIds !== undefined) {
            if (!Array.isArray(scope.siteIds) || scope.siteIds.some(id => !Number.isSafeInteger(Number(id)) || Number(id) <= 0)) throw new Error('加载站点范围无效。');
            scope.siteIds = [...new Set(scope.siteIds.map(id => str(Number(id))))];
        }
        return scope;
    }
    function scopeMatches(site, scope) {
        return ['manager', 'country'].every(key => scope[key + 'Mode'] === 'all' ||
            (scope[key + 'Mode'] === 'unassigned' ? !str(site[key]).trim() : str(site[key]).trim() === scope[key])) &&
            (!scope.siteId || str(site.id) === scope.siteId) && (scope.siteIds === undefined || scope.siteIds.includes(str(site.id)));
    }
    function scopeParams(scope) {
        const params = new URLSearchParams();
        ['manager', 'country'].forEach(key => {
            if (scope[key + 'Mode'] !== 'all') params.set(key + '_mode', scope[key + 'Mode']);
            if (scope[key + 'Mode'] === 'exact') params.set(key, scope[key]);
        });
        if (scope.siteId) params.set('site_id', scope.siteId);
        if (scope.siteIds !== undefined) params.set('site_ids', scope.siteIds.join(','));
        return params;
    }

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
            sku: str(row.sku), brands: list(row.brands), flavors: list(row.flavors),
            country: str(row.country === undefined ? site.country : row.country).trim()
        });
    }

    function matchesFilters(row, filters) {
        return (!filters.manager || valueKey(managerLabel(row.site)) === filters.manager) &&
            (!filters.country || countryKey(countryCode(row)) === filters.country) &&
            (!filters.site || str(row.site_id) === filters.site) &&
            (!filters.brand || (row.brands.length ? row.brands.some(value => valueKey(value) === filters.brand) : filters.brand === 'missing:brand')) &&
            (!filters.product || valueKey(row.product_name) === filters.product) &&
            (!filters.flavor || (row.flavors.length ? row.flavors.some(value => valueKey(value) === filters.flavor) : filters.flavor === 'missing:flavor')) &&
            (!filters.stock || stockValue(row) === filters.stock) &&
            (!filters.publish || str(row.status || 'unknown') === filters.publish) &&
            (!searchNormalized(filters.text) || searchNormalized([siteLabel(row.site), managerLabel(row.site), countryCode(row), row.product_name, row.name, row.sku, ...row.brands, ...row.flavors,
                ...(Array.isArray(row.attributes) ? row.attributes.flatMap(attribute => list(attribute.values)) : [])].join(' ')).includes(searchNormalized(filters.text)));
    }

    function facetOptions(rows, sites) {
        const maps = Object.fromEntries(['manager', 'country', 'site', 'brand', 'product', 'flavor', 'stock', 'publish'].map(key => [key, new Map()]));
        sites.forEach(site => {
            maps.manager.set(valueKey(managerLabel(site)), managerLabel(site));
            maps.site.set(str(site.id), siteLabel(site));
            maps.country.set(countryKey(site.country), str(site.country).trim() || '未设置国家');
        });
        rows.forEach(row => {
            maps.country.set(countryKey(countryCode(row)), countryCode(row) || '未设置国家');
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
        const parentConcurrency = Math.max(1, Math.min(2, Number(options.parentConcurrency) || 2));
        const emitInterval = options.emitInterval === undefined ? 100 : Math.max(0, Number(options.emitInterval) || 0);
        let epoch = 0, aborter = null, metadataEpoch = 0, metadataAborter = null;
        let dataVersion = 0, filtersVersion = 0, cachedRowsVersion = -1, cachedFacetVersion = -1;
        let cachedFilterVersion = -1, cachedFilterDataVersion = -1, cachedRows = [], cachedFiltered = [], cachedFacets = {};
        let emitTimer = null;
        const state = {phase: 'idle', queryMode: 'flavor', brand: '', keyword: '', sites: [], siteStates: new Map(), rows: new Map(), filters: emptyFilters(), page: 1, error: '',
            availableSites: [], metadataPhase: 'idle', metadataError: '', loadScope: emptyScope(), loadedScope: null, loadedAt: null};
        function parseSites(data) {
            if (!Array.isArray(data.sites)) throw new Error('站点列表格式无效，无法确认查询范围。');
            const ids = new Set();
            return data.sites.map(site => {
                const id = Number(site.id);
                if (!Number.isSafeInteger(id) || id <= 0 || ids.has(id)) throw new Error('站点列表有无效或重复编号，无法确认查询范围。');
                ids.add(id);
                return Object.assign({}, site, {id: str(id)});
            });
        }
        function snapshot() {
            if (cachedRowsVersion !== dataVersion) { cachedRows = [...state.rows.values()]; cachedRowsVersion = dataVersion; }
            const rows = cachedRows;
            if (cachedFilterDataVersion !== dataVersion || cachedFilterVersion !== filtersVersion) {
                cachedFiltered = rows.filter(row => matchesFilters(row, state.filters));
                cachedFilterDataVersion = dataVersion; cachedFilterVersion = filtersVersion;
            }
            const filtered = cachedFiltered;
            if (cachedFacetVersion !== dataVersion) { cachedFacets = facetOptions(rows, state.sites); cachedFacetVersion = dataVersion; }
            const siteStates = state.sites.map(site => Object.assign({}, state.siteStates.get(str(site.id))));
            const complete = siteStates.filter(site => site.status === 'complete').length;
            const failed = siteStates.filter(site => site.status === 'failed').length;
            const incomplete = siteStates.filter(site => ['failed', 'incomplete', 'stopped'].includes(site.status)).length;
            const pages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
            state.page = Math.max(1, Math.min(state.page, pages));
            const selectedSites = state.availableSites.filter(site => scopeMatches(site, state.loadScope));
            return {phase: state.phase, queryMode: state.queryMode, mode: state.queryMode, brand: state.brand, keyword: state.keyword, sites: state.sites.slice(), siteStates,
                availableSites: state.availableSites.slice(), metadataPhase: state.metadataPhase, metadataError: state.metadataError,
                loadScope: Object.assign({}, state.loadScope, state.loadScope.siteIds === undefined ? {} : {siteIds: state.loadScope.siteIds.slice()}), selectedSites, scopeCount: selectedSites.length,
                loadedScope: state.loadedScope ? Object.assign({}, state.loadedScope, {siteIds: state.loadedScope.siteIds.slice()}) : null, loadedAt: state.loadedAt,
                rows, filtered, filters: Object.assign({}, state.filters), options: cachedFacets, dataVersion,
                page: state.page, pages, pageRows: filtered.slice((state.page - 1) * PAGE_SIZE, state.page * PAGE_SIZE),
                counts: {matched: rows.length, filtered: filtered.length, totalSites: state.sites.length, completeSites: complete, failedSites: failed, incompleteSites: incomplete}, error: state.error,
                busy: state.phase === 'discovering' || state.phase === 'loading'};
        }
        function emit(immediate) {
            if (!options.onChange) return;
            if (immediate || !emitInterval) {
                if (emitTimer !== null) clearTimeout(emitTimer);
                emitTimer = null;
                options.onChange(snapshot());
            } else if (emitTimer === null) {
                emitTimer = setTimeout(() => { emitTimer = null; options.onChange(snapshot()); }, emitInterval);
            }
        }
        function current(run) { return run === epoch && aborter && !aborter.signal.aborted; }
        function begin() {
            if (aborter) aborter.abort();
            state.siteStates.forEach(site => { if (site.status === 'loading' || site.status === 'pending') site.status = 'stopped'; });
            if (emitTimer !== null) clearTimeout(emitTimer);
            emitTimer = null;
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
            emit(true);
        }
        async function discoverSites() {
            if (metadataAborter) metadataAborter.abort();
            metadataAborter = new AbortController();
            const run = ++metadataEpoch;
            state.metadataPhase = 'loading'; state.metadataError = ''; emit(true);
            try {
                const data = await readJson(fetcher, '/api/product-manager/catalog-sites', metadataAborter.signal, options.requestTimeout);
                if (run !== metadataEpoch) return;
                state.availableSites = parseSites(data);
                state.metadataPhase = 'ready';
            } catch (error) {
                if (run !== metadataEpoch) return;
                state.metadataPhase = 'error'; state.metadataError = error.message || str(error);
            }
            emit(true);
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
                if (state.loadedScope) scopeParams(state.loadedScope).forEach((value, key) => {
                    // site_id is the page's actual site, while site_ids bounds the full run.
                    if (key !== 'site_id') params.set(key, value);
                });
                if (parentId && parents.get(parentId).parent_token) params.set('parent_token', str(parents.get(parentId).parent_token));
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
                incoming.forEach(row => {
                    if (!state.rows.has(row.key)) siteState.matched += 1;
                    state.rows.set(row.key, row);
                });
                if (incoming.length) dataVersion += 1;
                siteState.scanned += Number(data.scanned) || 0;
                list(data.warnings).forEach(warning => { if (!siteState.warnings.includes(warning)) siteState.warnings.push(warning); });
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
                const parentIds = [...parents.keys()], activeParents = new Map();
                let parentCursor = 0, parentError = null;
                async function parentWorker() {
                    while (current(run) && !parentError && parentCursor < parentIds.length) {
                        const parentId = parentIds[parentCursor++];
                        activeParents.set(parentId, str(parents.get(parentId).name || ('产品 ' + parentId)));
                        siteState.currentProduct = [...activeParents.values()].join(' / ');
                        try {
                            const scanned = await scanPages(siteState, parentId, run, parents);
                            if (!current(run)) return;
                            const expectedIds = parents.get(parentId).variation_ids;
                            if (!Array.isArray(expectedIds)) throw new Error('可变产品缺少完整变体编号，当前站点结果不完整。');
                            const expected = new Set(expectedIds.map(id => str(Number(id))));
                            // Woo parent children list only publish/private; drafts
                            // remain real editable rows without expanding that list.
                            const seen = new Set([...scanned.sourceStatuses.entries()].filter(([, status]) => status === 'publish' || status === 'private').map(([id]) => id));
                            if (expected.size !== expectedIds.length || [...expected].some(id => !Number.isSafeInteger(Number(id)) || Number(id) <= 0) ||
                                seen.size !== expected.size || [...expected].some(id => !seen.has(id))) {
                                throw new Error('读取期间产品变体目录发生变化或缺失，当前站点结果不完整，请刷新后重试。');
                            }
                            siteState.variableDone += 1;
                        } catch (error) { if (current(run) && !parentError) parentError = error; }
                        finally {
                            activeParents.delete(parentId);
                            if (current(run)) { siteState.currentProduct = [...activeParents.values()].join(' / '); emit(); }
                        }
                    }
                }
                await Promise.all(Array.from({length: Math.min(parentConcurrency, parentIds.length)}, parentWorker));
                if (parentError) throw parentError;
                if (!current(run)) return;
                siteState.currentProduct = '';
                siteState.status = siteState.warnings.length ? 'incomplete' : 'complete';
            } catch (error) {
                if (!current(run)) return;
                siteState.status = 'failed';
                siteState.error = error.message || str(error);
            }
            emit(true);
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
            emit(true);
        }
        async function load(keyword, loadOptions) {
            const queryMode = loadOptions && loadOptions.queryMode || 'flavor';
            const brand = queryMode === 'brand' ? str(loadOptions && loadOptions.brand).trim() : '';
            const search = queryMode === 'flavor' ? str(keyword).trim() : '';
            const previousPhase = state.phase;
            let requestedScope;
            function invalid(message) {
                if (!state.loadedScope) state.phase = 'error';
                else if (state.phase === 'discovering') state.phase = previousPhase === 'loading' || previousPhase === 'discovering' ? 'stopped' : previousPhase;
                state.error = state.loadedScope ? '本次未加载成功：' + message + ' 当前仍显示原加载范围的结果，原加载时间保持不变。' : message; emit(true);
            }
            if (!['flavor', 'brand', 'all'].includes(queryMode) || (queryMode === 'brand' && !brand)) return invalid(queryMode === 'brand' ? '请输入或选择一个品牌，再加载所选站点。' : '请选择有效的查询方式。');
            if (brand.length > 200 || search.length > 200) return invalid('品牌或口味关键词不能超过 200 个字符。');
            try { requestedScope = normalizeScope(loadOptions && loadOptions.scope || state.loadScope); }
            catch (error) { return invalid(error.message); }
            if (state.metadataPhase === 'ready') {
                requestedScope.siteIds = state.availableSites.filter(site => scopeMatches(site, requestedScope)).map(site => site.id);
                if (!requestedScope.siteIds.length) return invalid('当前加载范围没有符合条件的站点。请调整负责人、国家或站点后再加载；旧结果仍保留。');
            } else if (requestedScope.siteIds && !requestedScope.siteIds.length) return invalid('当前加载范围没有站点，请调整后再加载。');
            const run = begin();
            state.phase = 'discovering'; state.error = ''; emit(true);
            try {
                const params = scopeParams(requestedScope), query = params.toString();
                const data = await readJson(fetcher, '/api/product-manager/catalog-sites' + (query ? '?' + query : ''), aborter.signal, options.requestTimeout);
                if (!current(run)) return;
                const discovered = parseSites(data), selected = discovered.filter(site => scopeMatches(site, requestedScope));
                if (!selected.length && (query || state.metadataPhase === 'ready')) return invalid('当前加载范围没有可读取的站点，可能权限或负责人已变化。旧结果仍保留，请重新选择范围。');
                if (!query) { state.availableSites = discovered; state.metadataPhase = 'ready'; state.metadataError = ''; }
                state.queryMode = queryMode; state.brand = brand; state.keyword = search;
                state.rows.clear(); state.sites = selected; state.siteStates.clear(); state.page = 1;
                state.loadedScope = Object.assign({}, requestedScope, {siteIds: selected.map(site => site.id)});
                state.loadedAt = new Date().toISOString();
                if (!loadOptions || !loadOptions.preserveFilters) state.filters = emptyFilters();
                filtersVersion += 1;
                dataVersion += 1;
                state.sites.forEach(site => state.siteStates.set(str(site.id), newSiteState(site)));
                await runSites(state.sites, run);
            } catch (error) {
                if (!current(run)) return;
                invalid(error.message || str(error));
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
            dataVersion += 1;
            targets.forEach(site => state.siteStates.set(str(site.id), newSiteState(site)));
            await runSites(targets, run);
        }
        return {load, stop, retryIncomplete, discoverSites, getSnapshot: snapshot,
            setLoadScope(scope) { state.loadScope = normalizeScope(scope, state.loadScope); state.error = ''; emit(true); },
            loadQuery(query) { return load(query.keyword || query.search || '', {queryMode: query.queryMode || query.mode || 'flavor', brand: query.brand, scope: query.scope, preserveFilters: !!query.preserveFilters}); },
            refresh() { return load(state.keyword, {queryMode: state.queryMode, brand: state.brand, scope: state.loadedScope || state.loadScope, preserveFilters: true}); },
            updateRows(rows) {
                rows.forEach(row => {
                    const site = state.sites.find(site => str(site.id) === str(row.site_id));
                    if (!site) return;
                    const normalizedRow = normalizeRow(row, site);
                    if (state.rows.has(normalizedRow.key)) state.rows.set(normalizedRow.key, normalizedRow);
                });
                dataVersion += 1; emit(true);
            },
            setFilters(filters) { state.filters = Object.assign({}, state.filters, filters); filtersVersion += 1; state.page = 1; emit(true); },
            resetFilters() { state.filters = emptyFilters(); filtersVersion += 1; state.page = 1; emit(true); },
            setPage(page) { state.page = Number(page) || 1; emit(true); }};
    }

    function resultHtml(snapshot, editor) {
        if (!snapshot.pageRows.length) {
            let message;
            if (snapshot.phase === 'idle') message = '选择加载范围，再选择按品牌、按口味或全部产品，点击「加载所选站点」开始查询。';
            else if (snapshot.counts.matched) message = '已加载的匹配结果中没有符合当前筛选条件的产品。可清空筛选后查看。';
            else if (snapshot.busy) message = '正在逐站读取，暂未收到匹配结果，请查看读取进度。';
            else if (snapshot.phase === 'complete' && snapshot.counts.totalSites) message = '已完整读取本次所选站点，此次查询未匹配到产品或口味。';
            else if (snapshot.phase === 'complete') message = '当前账号没有可查询的站点，请联系管理员核对产品管理权限。';
            else message = '当前结果尚不完整，暂未匹配到产品。不能据此判断全部站点没有符合查询条件的产品，请重试异常站点或重新加载。';
            return '<div class="pm-empty-state">' + escapeHtml(message) + '</div>';
        }
        if (editor) return editor.tableHtml(snapshot);
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
        let latest, scheduled = false, editor = null, tableSignature = '';
        const filterOptions = new Map(), fetcher = mountOptions && mountOptions.fetch || global.fetch.bind(global);
        function scheduleRender(snapshot) {
            latest = snapshot;
            if (scheduled) return;
            scheduled = true;
            (global.requestAnimationFrame || (callback => setTimeout(callback, 0)))(() => { scheduled = false; render(latest); });
        }
        const controller = createController({fetch: fetcher, onChange: scheduleRender});
        const hasScopeControls = !!($('pmCatalogLoadManager') && $('pmCatalogLoadCountry') && $('pmCatalogLoadSite'));
        function syncSelect(element) {
            if (global.ProductManagerSelects && element) global.ProductManagerSelects.sync(element);
        }
        if (global.ProductManagerCatalogEditor && $('pmCatalogSelectedCount')) editor = global.ProductManagerCatalogEditor.createUI({document, root, catalog: controller, fetch: fetcher,
            onChange: () => scheduleRender(controller.getSnapshot())});
        function render(snapshot) {
            const editorBusy = editor && editor.getSnapshot().busy;
            $('pmCatalogLoad').disabled = !!editorBusy || (hasScopeControls && (snapshot.metadataPhase !== 'ready' || !snapshot.scopeCount));
            $('pmCatalogStop').disabled = !snapshot.busy;
            $('pmCatalogRetry').disabled = editorBusy || snapshot.busy || !snapshot.counts.incompleteSites;
            $('pmCatalogRefresh').disabled = editorBusy || !snapshot.loadedScope;
            let status = {idle: '选择加载范围和产品条件；加载后可在本页反复搜索、筛选。', discovering: '正在确认所选范围的站点权限…', loading: '正在读取所选站点的产品与全部变体，匹配结果会陆续显示。', complete: '查询完成，本次所选站点读取完整。', incomplete: '查询结束，存在不完整站点；已保留成功读取的匹配结果。', stopped: '已停止；已读取的匹配结果保留，未完成站点可单独重试。', error: snapshot.error}[snapshot.phase];
            if (snapshot.phase === 'complete' && !snapshot.counts.totalSites) status = '当前账号没有可查询的站点，请核对产品管理权限。';
            if (snapshot.phase === 'stopped' && !snapshot.counts.totalSites) status = '已停止确认站点范围，请重新加载以继续查询。';
            if (snapshot.phase !== 'idle' && snapshot.phase !== 'discovering') {
                const scope = snapshot.queryMode === 'brand' ? (snapshot.brand ? `品牌「${snapshot.brand}」` : '品牌查询') : snapshot.queryMode === 'all' || !snapshot.keyword ? '全部产品' : `口味 / 关键词「${snapshot.keyword}」`;
                status = scope + '：' + status;
            }
            $('pmCatalogStatus').textContent = snapshot.error || status;
            $('pmCatalogStatus').classList.toggle('text-warning', !!snapshot.error || ['incomplete', 'stopped', 'error'].includes(snapshot.phase));
            renderScope(snapshot);
            const counts = snapshot.counts;
            $('pmCatalogCounts').textContent = `已匹配 ${counts.matched} 条 · 筛选后 ${counts.filtered} 条 · 完整站点 ${counts.completeSites} / ${counts.totalSites} · 失败 ${counts.failedSites} · 不完整 ${counts.incompleteSites}`;
            $('pmCatalogProgress').innerHTML = snapshot.siteStates.map(site => {
                const problems = [site.error, ...site.warnings].filter(Boolean).map(escapeHtml).join('；');
                const progress = `产品 ${site.productPages}${Number.isFinite(Number(site.productTotalPages)) && site.productTotalPages !== null ? ' / ' + Number(site.productTotalPages) : ''} 页 · 已检查 ${site.scanned} 项 · 匹配 ${site.matched} 条 · 变体 ${site.variableDone} / ${site.variableTotal} 个产品（${site.variationPages} 页）`;
                return `<div class="border-bottom border-secondary border-opacity-25 py-2"><div class="d-flex justify-content-between gap-2"><strong>${escapeHtml(siteLabel(site.site))}</strong><span class="${site.status === 'complete' ? 'text-success' : ['failed', 'incomplete', 'stopped'].includes(site.status) ? 'text-warning' : 'text-info'}">${escapeHtml(SITE_LABELS[site.status])}</span></div><div class="small text-muted">${escapeHtml(managerLabel(site.site))} · ${escapeHtml(str(site.site.country).trim() || '未设置国家')} · ${escapeHtml(progress)}${site.currentProduct ? ' · 正在读取：' + escapeHtml(site.currentProduct) : ''}</div>${problems ? '<div class="small text-warning">' + problems + '</div>' : ''}</div>`;
            }).join('');
            Object.entries(FILTER_IDS).forEach(([key, id]) => {
                const element = $(id);
                if (!element) return;
                if (key === 'text') { if (element.value !== snapshot.filters.text) element.value = snapshot.filters.text; return; }
                const selected = snapshot.filters[key], sourceChoices = snapshot.options[key];
                const choices = sourceChoices.slice();
                if (selected && !choices.some(choice => choice.value === selected)) {
                    const old = [...element.options].find(option => option.value === selected);
                    choices.push({value: selected, label: old ? old.textContent : selected});
                }
                const previous = filterOptions.get(key), extra = selected && !sourceChoices.some(choice => choice.value === selected) ? selected : '';
                let changed = false;
                if (!previous || previous.source !== sourceChoices || previous.extra !== extra) {
                    const html = '<option value="">全部</option>' + choices.map(choice => `<option value="${escapeHtml(choice.value)}">${escapeHtml(choice.label)}</option>`).join('');
                    if (document.activeElement === element) element.pendingCatalogOptions = html;
                    else { element.innerHTML = html; element.pendingCatalogOptions = null; changed = true; }
                    filterOptions.set(key, {source: sourceChoices, extra});
                }
                if (element.value !== selected && [...element.options].some(option => option.value === selected)) { element.value = selected; changed = true; }
                if (changed) syncSelect(element);
            });
            const signature = JSON.stringify([snapshot.dataVersion, snapshot.phase, snapshot.page, snapshot.filters, editor ? editor.tableVersion : 0]);
            if (tableSignature !== signature) { $('pmCatalogResults').innerHTML = resultHtml(snapshot, editor); tableSignature = signature; }
            if (editor) editor.render(snapshot);
        }
        function scopeValue(scope, key) { return scope[key + 'Mode'] === 'all' ? '' : scope[key + 'Mode'] === 'unassigned' ? 'unassigned' : 'exact:' + scope[key]; }
        function renderScope(snapshot) {
            if (!hasScopeControls) return;
            ['manager', 'country'].forEach(key => {
                const element = $(key === 'manager' ? 'pmCatalogLoadManager' : 'pmCatalogLoadCountry');
                const values = [...new Set(snapshot.availableSites.map(site => str(site[key]).trim()))].sort((a, b) => a.localeCompare(b, 'zh-CN'));
                const choices = values.map(value => ({value: value ? 'exact:' + value : 'unassigned', label: value || (key === 'manager' ? '未分配负责人' : '未设置国家')}));
                const selected = scopeValue(snapshot.loadScope, key);
                if (selected && !choices.some(choice => choice.value === selected)) choices.push({value: selected, label: (snapshot.loadScope[key] || (key === 'manager' ? '未分配负责人' : '未设置国家')) + '（当前无可选站点）'});
                const html = `<option value="">全部${key === 'manager' ? '负责人' : '国家'}</option>` + choices.map(choice => `<option value="${escapeHtml(choice.value)}">${escapeHtml(choice.label)}</option>`).join('');
                let changed = false;
                if (element.catalogScopeOptions !== html) { element.innerHTML = html; element.catalogScopeOptions = html; changed = true; }
                if (element.value !== selected) { element.value = selected; changed = true; }
                if (element.disabled !== (snapshot.metadataPhase !== 'ready')) { element.disabled = snapshot.metadataPhase !== 'ready'; changed = true; }
                if (changed) syncSelect(element);
            });
            const siteElement = $('pmCatalogLoadSite');
            const siteHtml = snapshot.availableSites.map(site => `<option value="${escapeHtml(site.id)}">${escapeHtml(siteLabel(site))} · ${escapeHtml(managerLabel(site))} · ${escapeHtml(str(site.country).trim() || '未设置国家')}</option>`).join('');
            let sitesChanged = false;
            if (!siteElement.multiple) { siteElement.multiple = true; sitesChanged = true; }
            if (siteElement.catalogScopeOptions !== siteHtml) { siteElement.innerHTML = siteHtml; siteElement.catalogScopeOptions = siteHtml; sitesChanged = true; }
            const chosenIds = snapshot.loadScope.siteIds === undefined ? null : new Set(snapshot.loadScope.siteIds);
            [...siteElement.options].forEach(option => {
                const selected = (!snapshot.loadScope.siteId || option.value === snapshot.loadScope.siteId) && (chosenIds === null || chosenIds.has(option.value));
                if (option.selected !== selected) { option.selected = selected; sitesChanged = true; }
            });
            if (siteElement.disabled !== (snapshot.metadataPhase !== 'ready')) { siteElement.disabled = snapshot.metadataPhase !== 'ready'; sitesChanged = true; }
            if (sitesChanged) syncSelect(siteElement);
            if ($('pmCatalogScopeCount')) $('pmCatalogScopeCount').textContent = snapshot.metadataPhase === 'ready' ? `本次将加载 ${snapshot.scopeCount} / ${snapshot.availableSites.length} 个有权限站点` : snapshot.metadataPhase === 'error' ? '站点列表读取失败' : '正在获取可选站点（不读取网站商品）…';
            if ($('pmCatalogReloadSites')) $('pmCatalogReloadSites').disabled = snapshot.metadataPhase === 'loading';
            const changed = snapshot.loadedScope && (['manager', 'managerMode', 'country', 'countryMode', 'siteId'].some(key => snapshot.loadScope[key] !== snapshot.loadedScope[key]) || snapshot.selectedSites.map(site => site.id).join(',') !== snapshot.loadedScope.siteIds.join(','));
            if ($('pmCatalogScopeStatus')) {
                $('pmCatalogScopeStatus').textContent = snapshot.metadataError ? '站点列表未更新：' + snapshot.metadataError + (snapshot.loadedScope ? ' 原加载范围、结果和加载时间保持不变。' : ' 请点击「刷新站点列表」重试。') : (changed ? '待加载范围已修改；当前结果仍来自下方标明的原范围，点击「加载所选站点」才读取新范围。' : '调整范围只影响下一次加载，不会清除已有结果。只有加载、重新加载和重试会读取网站商品。');
                $('pmCatalogScopeStatus').classList.toggle('text-warning', !!snapshot.metadataError);
            }
            if ($('pmCatalogLoadedScope') && snapshot.loadedScope) {
                const scope = snapshot.loadedScope;
                const manager = scope.managerMode === 'all' ? '全部负责人' : scope.managerMode === 'unassigned' ? '未分配负责人' : scope.manager;
                const country = scope.countryMode === 'all' ? '全部国家' : scope.countryMode === 'unassigned' ? '未设置国家' : scope.country;
                const names = snapshot.sites.slice(0, 2).map(siteLabel).join('、');
                const site = '本次勾选站点' + (names ? `：${names}${scope.siteIds.length > 2 ? '等' : ''}` : '');
                const query = snapshot.queryMode === 'brand' ? `品牌「${snapshot.brand}」` : snapshot.queryMode === 'flavor' && snapshot.keyword ? `口味 / 关键词「${snapshot.keyword}」` : '全部产品';
                $('pmCatalogLoadedScope').textContent = `本页已加载范围：${manager} · ${country} · ${site}（${scope.siteIds.length} 个站点）· ${query}。加载时间：${new Date(snapshot.loadedAt).toLocaleString()}。结果筛选仅在此范围内进行，刷新浏览器后需重新加载。`;
            }
        }
        function syncQueryForm() {
            const mode = $('pmCatalogQueryMode') ? $('pmCatalogQueryMode').value : 'flavor';
            if ($('pmCatalogFlavorInputWrap')) $('pmCatalogFlavorInputWrap').classList.toggle('d-none', mode !== 'flavor');
            if ($('pmCatalogBrandInputWrap')) $('pmCatalogBrandInputWrap').classList.toggle('d-none', mode !== 'brand');
        }
        function loadFormQuery() {
            if (editor && !editor.canReload()) return;
            const snapshot = controller.getSnapshot();
            if (hasScopeControls && (snapshot.metadataPhase !== 'ready' || !snapshot.scopeCount)) return;
            const queryMode = $('pmCatalogQueryMode') ? $('pmCatalogQueryMode').value : 'flavor';
            return controller.load($('pmCatalogSearch').value, {queryMode, brand: $('pmCatalogBrandInput') ? $('pmCatalogBrandInput').value : ''});
        }
        function restoreLoadedQueryForm() {
            const snapshot = controller.getSnapshot();
            if ($('pmCatalogQueryMode')) $('pmCatalogQueryMode').value = snapshot.queryMode;
            syncSelect($('pmCatalogQueryMode'));
            if ($('pmCatalogBrandInput')) $('pmCatalogBrandInput').value = snapshot.brand;
            $('pmCatalogSearch').value = snapshot.keyword;
            syncQueryForm();
        }
        $('pmCatalogLoad').addEventListener('click', loadFormQuery);
        [$('pmCatalogSearch'), $('pmCatalogBrandInput')].filter(Boolean).forEach(input => input.addEventListener('keydown', event => {
            if (event.key === 'Enter') { event.preventDefault(); loadFormQuery(); }
        }));
        if ($('pmCatalogQueryMode')) $('pmCatalogQueryMode').addEventListener('change', syncQueryForm);
        if (hasScopeControls) {
            function scopeFromForm() {
                const scope = {siteId: '', siteIds: [...$('pmCatalogLoadSite').selectedOptions].map(option => option.value)};
                ['manager', 'country'].forEach(key => {
                    const value = $(key === 'manager' ? 'pmCatalogLoadManager' : 'pmCatalogLoadCountry').value;
                    scope[key + 'Mode'] = value === '' ? 'all' : value === 'unassigned' ? 'unassigned' : 'exact';
                    scope[key] = value.startsWith('exact:') ? value.slice(6) : '';
                });
                controller.setLoadScope(scope);
            }
            ['pmCatalogLoadManager', 'pmCatalogLoadCountry', 'pmCatalogLoadSite'].forEach(id => $(id).addEventListener('change', scopeFromForm));
            if ($('pmCatalogReloadSites')) $('pmCatalogReloadSites').addEventListener('click', controller.discoverSites);
        }
        $('pmCatalogRefresh').addEventListener('click', () => { if (editor && !editor.canReload()) return; restoreLoadedQueryForm(); controller.refresh(); });
        $('pmCatalogStop').addEventListener('click', controller.stop);
        $('pmCatalogRetry').addEventListener('click', () => { restoreLoadedQueryForm(); controller.retryIncomplete(); });
        $('pmCatalogResetFilters').addEventListener('click', controller.resetFilters);
        Object.entries(FILTER_IDS).forEach(([key, id]) => {
            if (!$(id)) return;
            $(id).addEventListener(key === 'text' ? 'input' : 'change', event => controller.setFilters({[key]: event.target.value}));
            if (key !== 'text') $(id).addEventListener('blur', () => {
                const element = $(id);
                if (element.pendingCatalogOptions) { element.innerHTML = element.pendingCatalogOptions; element.pendingCatalogOptions = null; element.value = controller.getSnapshot().filters[key]; syncSelect(element); }
            });
        });
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
        root.productManagerCatalogEditor = editor;
        if (hasScopeControls) controller.discoverSites();
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
