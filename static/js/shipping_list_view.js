(function (root, factory) {
    if (typeof module === 'object' && module.exports) {
        module.exports = factory();
    } else {
        root.ShippingListView = factory();
    }
}(typeof self !== 'undefined' ? self : this, function () {
    'use strict';

    function paginate(orders, requestedPage, requestedSize) {
        const all = Array.isArray(orders) ? orders : [];
        const size = Number(requestedSize);
        const perPage = Number.isInteger(size) && size > 0 ? Math.min(size, 50) : 50;
        const pages = Math.max(1, Math.ceil(all.length / perPage));
        const requested = Number(requestedPage);
        const page = Math.min(pages, Number.isInteger(requested) && requested > 0 ? requested : 1);
        return {
            orders: all.slice((page - 1) * perPage, page * perPage),
            total: all.length,
            page: page,
            pages: pages,
            per_page: perPage,
            has_previous: page > 1,
            has_next: page < pages
        };
    }

    return { paginate: paginate };
}));
