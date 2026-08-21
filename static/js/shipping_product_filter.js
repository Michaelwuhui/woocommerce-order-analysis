(function (root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) {
        module.exports = api;
    }
    root.ShippingProductFilter = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';

    function normalize(value) {
        return String(value == null ? '' : value)
            .normalize('NFKC')
            .trim()
            .toLowerCase();
    }

    function quantity(value) {
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function productOptions(orders) {
        const products = new Map();

        (orders || []).forEach((order, orderIndex) => {
            const orderKey = String(order && (order.id || order.number) || orderIndex);
            (order && order.products || []).forEach(product => {
                const name = String(product && product.name || '').trim();
                const key = normalize(name);
                if (!key) return;

                if (!products.has(key)) {
                    products.set(key, {
                        name: name,
                        units: 0,
                        orderKeys: new Set()
                    });
                }
                const item = products.get(key);
                item.units += quantity(product.quantity);
                item.orderKeys.add(orderKey);
            });
        });

        return Array.from(products.values())
            .map(item => ({
                name: item.name,
                units: item.units,
                orderCount: item.orderKeys.size
            }))
            .sort((left, right) => (
                right.units - left.units
                || left.name.localeCompare(right.name, 'zh-CN', { sensitivity: 'base' })
            ));
    }

    function searchOptions(options, query, limit) {
        const normalizedQuery = normalize(query);
        const maxResults = Number.isFinite(Number(limit)) ? Math.max(1, Number(limit)) : 8;
        if (normalizedQuery.length < 2) return [];

        const tokens = normalizedQuery.split(/\s+/).filter(Boolean);
        return (Array.isArray(options) ? options : [])
            .map(option => {
                const normalizedName = normalize(option && option.name);
                if (!normalizedName || !tokens.every(token => normalizedName.includes(token))) {
                    return null;
                }
                let score = 3;
                if (normalizedName === normalizedQuery) score = 0;
                else if (normalizedName.startsWith(normalizedQuery)) score = 1;
                else if (normalizedName.includes(normalizedQuery)) score = 2;
                return { option: option, score: score };
            })
            .filter(Boolean)
            .sort((left, right) => (
                left.score - right.score
                || Number(right.option.units || 0) - Number(left.option.units || 0)
                || String(left.option.name || '').localeCompare(
                    String(right.option.name || ''), 'zh-CN', { sensitivity: 'base' }
                )
            ))
            .slice(0, maxResults)
            .map(item => item.option);
    }

    function summarize(orders, query) {
        const allOrders = Array.isArray(orders) ? orders : [];
        const normalizedQuery = normalize(query);
        const options = productOptions(allOrders);
        const exactMatch = normalizedQuery
            ? options.some(item => normalize(item.name) === normalizedQuery)
            : false;

        let units = 0;
        const visibleOrders = [];

        allOrders.forEach(order => {
            const products = Array.isArray(order && order.products) ? order.products : [];
            const matchedProducts = normalizedQuery
                ? products.filter(product => {
                    const name = normalize(product && product.name);
                    return exactMatch ? name === normalizedQuery : name.includes(normalizedQuery);
                })
                : products;

            if (!normalizedQuery || matchedProducts.length) {
                visibleOrders.push(order);
                matchedProducts.forEach(product => {
                    units += quantity(product.quantity);
                });
            }
        });

        return {
            orders: visibleOrders,
            orderCount: visibleOrders.length,
            units: units,
            productCount: options.length,
            exactMatch: exactMatch,
            query: String(query == null ? '' : query).trim()
        };
    }

    return {
        normalize: normalize,
        productOptions: productOptions,
        searchOptions: searchOptions,
        summarize: summarize
    };
}));
