(function (global) {
    'use strict';

    const text = value => value === null || value === undefined ? '' : String(value);
    const escapeHtml = value => text(value).replace(/[&<>"']/g, character => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[character]));
    const number = value => Number.isFinite(Number(value)) ? Number(value) : 0;
    const array = value => Array.isArray(value) ? value : [];
    const money = value => number(value).toLocaleString('en-US', {
        minimumFractionDigits: 2, maximumFractionDigits: 2
    });
    const searchCache = new WeakMap();

    function searchText(row) {
        if (!searchCache.has(row)) {
            const emails = array(row.identity_emails);
            searchCache.set(row, [text(row.name), (emails.length ? emails : [row.email]).map(text).join(' '),
                text(row.phone)].join(' ').toLowerCase());
        }
        return searchCache.get(row);
    }

    function matchesCustomer(row, filters) {
        if (filters.lossOnly && !(number(row.total_loss) > 0)) return false;
        if (filters.multiSiteOnly && !(number(row.site_count || 1) >= 2)) return false;
        if (filters.mergedOnly && !(number(row.identity_email_count || 1) >= 2)) return false;
        return !filters.query || searchText(row).indexOf(filters.query) !== -1;
    }

    function sourceText(row, config) {
        const source = text(row.source);
        const manager = text((config.siteManagers || {})[source]);
        const short = source.replace('https://www.', '').replace('https://', '');
        return [manager, config.sourceDisplayMode === 'manager_only' ? '' : short].filter(Boolean).join(' ');
    }

    // Orthogonal values keep sorting/filtering independent of display HTML.
    // DataTables only asks for display HTML when it creates a visible row.
    function sortValue(column, row, config) {
        switch (column) {
            case 0: return text(row.name || 'Unknown');
            case 1: return sourceText(row, config);
            case 2: return number(row.site_count || 1);
            case 3: return text(row.email) + ' ' + text(row.phone);
            case 4: return text(row.tier || '新客');
            case 5: return number(row.successful_orders);
            case 6: return number(row.total_spent_cny);
            case 7: return number(row.total_loss);
            case 8: return text(row.last_order_date).replace('T', ' ').slice(0, 10);
            case 9: return array(row.actions).map(action => text(action.text)).join(' ');
            default: return '';
        }
    }

    function nameHtml(row) {
        const name = text(row.name);
        const tierColor = {VIP: 'warning', '优质': 'success', '普通': 'primary'}[row.tier] || 'secondary';
        let html = '<div class="d-flex align-items-center"><div class="avatar-circle me-2 bg-gradient-' + tierColor + '">'
            + escapeHtml(name ? Array.from(name)[0].toUpperCase() : '?') + '</div><div>'
            + '<div class="fw-bold text-white">' + escapeHtml(name || 'Unknown') + '</div>'
            + '<small class="text-white-50">' + escapeHtml(row.email) + '</small>';
        const count = number(row.identity_email_count || 1);
        if (count > 1) {
            const matchedMap = {phone: '同手机', address: '同地址'};
            const reasons = array(row.identity_matched_by).map(reason => matchedMap[reason] || text(reason)).join('+');
            const tooltip = '同一身份合并了 ' + count + ' 个邮箱（依据：' + reasons + '）：\n'
                + array(row.identity_emails).map(email => '· ' + text(email)).join('\n');
            html += '<div class="mt-1"><span class="badge" style="background:rgba(167,139,250,0.18);color:#a78bfa;border:1px solid rgba(167,139,250,0.5);cursor:help;" title="'
                + escapeHtml(tooltip) + '"><i class="bi bi-people-fill me-1"></i>合并 ' + count + ' 邮箱</span></div>';
        }
        return html + '</div></div>';
    }

    function sourceHtml(row, config) {
        const source = text(row.source);
        const manager = text((config.siteManagers || {})[source]);
        return (manager ? '<span class="badge me-1" style="background:rgba(23,162,184,0.3);color:#17a2b8;">' + escapeHtml(manager) + '</span>' : '')
            + (config.sourceDisplayMode === 'manager_only' ? '' : '<span class="badge bg-dark border border-secondary">'
                + escapeHtml(source.replace('https://www.', '').replace('https://', '')) + '</span>');
    }

    function sitesHtml(row) {
        const count = number(row.site_count || 1);
        if (count < 2) return '<span class="text-secondary">—</span>';
        const tooltip = '该客户在 ' + count + ' 个站点下过单：\n' + array(row.site_list).map(site => '· '
            + text(site.short) + (site.manager ? ' (' + text(site.manager) + ')' : '')).join('\n');
        return '<span class="badge" style="background:rgba(93,208,230,0.18);color:#5dd0e6;border:1px solid rgba(93,208,230,0.5);cursor:help;" title="'
            + escapeHtml(tooltip) + '"><i class="bi bi-diagram-3 me-1"></i>' + count + ' 站</span>';
    }

    function contactHtml(row) {
        return '<div class="btn-group"><a href="mailto:' + escapeHtml(row.email)
            + '" class="btn btn-sm btn-outline-secondary" title="发送邮件"><i class="bi bi-envelope"></i></a>'
            + (row.phone ? '<a href="tel:' + escapeHtml(row.phone)
                + '" class="btn btn-sm btn-outline-secondary" title="拨打电话"><i class="bi bi-telephone"></i></a>' : '') + '</div>';
    }

    function tierHtml(row) {
        const tiers = {VIP: ['warning text-dark', 'star-fill'], '优质': ['success', 'gem'],
            '普通': ['primary', 'person-check'], '新客': ['secondary', 'person']};
        const tier = Object.prototype.hasOwnProperty.call(tiers, row.tier) ? row.tier : '新客';
        return '<span class="badge bg-' + tiers[tier][0] + '"><i class="bi bi-' + tiers[tier][1] + ' me-1"></i>' + tier + '</span>';
    }

    function spendingHtml(row) {
        const missing = array(row.missing_exchange_rates);
        return '¥' + money(row.total_spent_cny) + (missing.length
            ? '<i class="bi bi-exclamation-triangle-fill text-warning ms-1" title="'
                + escapeHtml('以下汇率缺失，人民币金额未完全折算：' + missing.map(text).join(', ')) + '"></i>' : '');
    }

    function lossHtml(row) {
        const undelivered = number(row.undelivered_orders), returns = number(row.problem_return_orders);
        if (!(undelivered > 0 || returns > 0)) return '<span class="text-secondary">—</span>';
        let html = '';
        if (undelivered > 0) {
            const tooltip = '拒收 ' + undelivered + ' 单 / 共 ' + number(row.total_orders) + ' 单 = ' + number(row.refusal_rate) + '%';
            html += '<span class="badge" style="background:rgba(168,85,247,0.2);color:#c084fc;border:1px solid rgba(168,85,247,0.4);" title="'
                + escapeHtml(tooltip) + '">未送达 ' + undelivered
                + (number(row.refusal_rate) >= 30 ? '<i class="bi bi-exclamation-triangle-fill ms-1" style="color:#fbbf24;" title="高拒收率"></i>' : '') + '</span>';
        }
        if (returns > 0) html += '<span class="badge ms-1" style="background:rgba(220,53,69,0.2);color:#f87171;border:1px solid rgba(220,53,69,0.4);" title="问题退货（调包/少件/损坏）'
            + returns + ' 单">问题退货 ' + returns + '</span>';
        if (number(row.shipping_loss_total) > 0) html += '<div><small class="text-warning">运费损失 '
            + money(row.shipping_loss_total) + ' ' + escapeHtml(row.currency) + '</small></div>';
        if (number(row.product_loss_total) > 0) html += '<div><small class="text-warning">货值损失 '
            + money(row.product_loss_total) + ' ' + escapeHtml(row.currency) + '</small></div>';
        return html;
    }

    function actionsHtml(row) {
        const allowedTypes = new Set(['success', 'primary', 'warning', 'info', 'danger', 'secondary', 'dark', 'light']);
        return array(row.actions).map(action => {
            const type = allowedTypes.has(action.type) ? action.type : 'secondary';
            const icon = /^[a-z0-9-]+$/.test(text(action.icon)) ? action.icon : 'info-circle';
            return '<span class="badge rounded-pill bg-' + type + ' bg-opacity-25 text-' + type
                + ' border border-' + type + ' me-1 mb-1" title="' + escapeHtml(action.text)
                + '"><i class="bi bi-' + icon + ' me-1"></i>' + escapeHtml(action.text) + '</span>';
        }).join('');
    }

    function renderCell(column, type, row, config) {
        config = config || {};
        if (type !== 'display') return sortValue(column, row, config);
        switch (column) {
            case 0: return nameHtml(row);
            case 1: return sourceHtml(row, config);
            case 2: return sitesHtml(row);
            case 3: return contactHtml(row);
            case 4: return tierHtml(row);
            case 5: return String(number(row.successful_orders));
            case 6: return spendingHtml(row);
            case 7: return lossHtml(row);
            case 8: return escapeHtml(sortValue(column, row, config));
            case 9: return actionsHtml(row);
            case 10: return '<button type="button" class="btn btn-sm btn-outline-info" data-customer-detail="'
                + escapeHtml(row.email) + '">详情</button>';
            default: return '';
        }
    }

    function buildColumns(config) {
        return Array.from({length: 11}, (_, column) => ({
            data: null,
            render: function (_data, type, row) { return renderCell(column, type, row, config); },
            className: column === 6 ? 'text-end fw-bold text-success' : [2, 5, 7].includes(column) ? 'text-center' : '',
            type: [2, 5, 6, 7].includes(column) ? 'num' : 'string',
            orderable: ![3, 9, 10].includes(column)
        }));
    }

    function mount(document, jquery, config, rows) {
        const table = document.getElementById('customersTable');
        if (!table || !jquery || !jquery.fn.dataTable) return null;
        const filters = {lossOnly: false, multiSiteOnly: false, mergedOnly: false, query: ''};
        jquery.fn.dataTable.ext.search.push(function (settings, _data, dataIndex, rowData) {
            if (settings.nTable !== table) return true;
            const row = rowData || settings.aoData[dataIndex]._aData;
            return matchesCustomer(row, filters);
        });
        const dt = jquery(table).DataTable({
            data: rows,
            deferRender: true,
            // Otherwise width measurement asks for display HTML for every row.
            autoWidth: false,
            columns: buildColumns(config),
            order: [[6, 'desc']],
            dom: 'lrtip',
            pageLength: 25,
            language: {
                sProcessing: '处理中...', sLengthMenu: '显示 _MENU_ 项结果', sZeroRecords: '没有匹配结果',
                sInfo: '显示第 _START_ 至 _END_ 项结果，共 _TOTAL_ 项', sInfoEmpty: '显示第 0 至 0 项结果，共 0 项',
                sInfoFiltered: '(由 _MAX_ 项结果过滤)', sInfoPostFix: '', sSearch: '搜索:', sUrl: '',
                sEmptyTable: '表中数据为空', sLoadingRecords: '载入中...', sInfoThousands: ',',
                oPaginate: {sFirst: '首页', sPrevious: '上页', sNext: '下页', sLast: '末页'},
                oAria: {sSortAscending: ': 以升序排列此列', sSortDescending: ': 以降序排列此列'}
            }
        });
        let searchTimer;
        const searchBox = document.getElementById('customerSearchBox');
        function applyQuery() {
            if (searchTimer) global.clearTimeout(searchTimer);
            filters.query = text(searchBox && searchBox.value).trim().toLowerCase();
        }
        if (searchBox) searchBox.addEventListener('input', function () {
            if (searchTimer) global.clearTimeout(searchTimer);
            searchTimer = global.setTimeout(function () { applyQuery(); dt.draw(); }, 300);
        });
        function bindToggle(id, key, active, inactive, order) {
            const button = document.getElementById(id);
            if (!button) return;
            button.addEventListener('click', function () {
                filters[key] = !filters[key];
                applyQuery();
                button.classList.toggle('active', filters[key]);
                if (key === 'mergedOnly') {
                    button.style.background = filters[key] ? '#a78bfa' : '';
                    button.style.color = filters[key] ? '#fff' : '#a78bfa';
                } else {
                    const color = key === 'lossOnly' ? 'danger' : 'info';
                    button.classList.toggle('btn-' + color, filters[key]);
                    button.classList.toggle('btn-outline-' + color, !filters[key]);
                }
                button.innerHTML = filters[key] ? active : inactive;
                if (order) dt.order([filters[key] ? order : 6, 'desc']);
                dt.draw();
            });
        }
        bindToggle('toggleLossOnlyBtn', 'lossOnly', '<i class="bi bi-funnel-fill me-1"></i>显示全部客户',
            '<i class="bi bi-funnel me-1"></i>只看有损失的客户', 7);
        bindToggle('toggleMultiSiteBtn', 'multiSiteOnly', '<i class="bi bi-funnel-fill me-1"></i>显示全部客户',
            '<i class="bi bi-diagram-3 me-1"></i>只看多站客户', 2);
        bindToggle('toggleMergedBtn', 'mergedOnly', '<i class="bi bi-funnel-fill me-1"></i>显示全部客户',
            '<i class="bi bi-people-fill me-1"></i>只看聚合客户');
        table.addEventListener('click', function (event) {
            const button = event.target.closest('[data-customer-detail]');
            if (button && table.contains(button) && typeof global.showCustomerDetail === 'function') {
                global.showCustomerDetail(button.getAttribute('data-customer-detail'));
            }
        });
        return dt;
    }

    const api = {escapeHtml, searchText, matchesCustomer, sortValue, renderCell, buildColumns, mount};
    if (typeof module === 'object' && module.exports) module.exports = api;
    global.CustomerTable = api;
    if (global.document) {
        const start = function () {
            const payload = global.document.getElementById('customersTableData');
            const config = global.document.getElementById('customersTableConfig');
            if (payload && config) mount(global.document, global.jQuery, JSON.parse(config.textContent), JSON.parse(payload.textContent));
        };
        if (global.document.readyState === 'loading') global.document.addEventListener('DOMContentLoaded', start);
        else start();
    }
})(typeof window !== 'undefined' ? window : globalThis);
