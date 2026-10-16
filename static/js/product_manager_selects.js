(function (global) {
    'use strict';

    const text = value => value == null ? '' : String(value);
    const normalizeSearch = value => text(value).toLocaleLowerCase()
        .replace(/[łøđðþæœß]/g, char => ({ł: 'l', ø: 'o', đ: 'd', ð: 'd', þ: 'th', æ: 'ae', œ: 'oe', ß: 'ss'}[char]))
        .normalize('NFKD').replace(/\p{M}/gu, '').replace(/\s+/g, ' ').trim();

    function create(document) {
        const win = document.defaultView || global;
        const models = new Map(), excluded = new WeakSet(), pending = new Set();
        let serial = 0, frame = null, positionFrame = null, current = null, panel = null;
        let search, listbox, status, allButton, clearButton, activeOption = null, visible = [];
        let observing = false, removed = false;
        const raf = callback => (win.requestAnimationFrame || win.setTimeout)(callback);
        const isDisabled = select => select.disabled || select.matches(':disabled');
        const optionDisabled = option => option.disabled || (option.parentElement.tagName === 'OPTGROUP' && option.parentElement.disabled);
        const optionHidden = option => option.hidden || (option.parentElement.tagName === 'OPTGROUP' && option.parentElement.hidden);
        const attr = (element, name, value) => { if (element.getAttribute(name) !== text(value)) element.setAttribute(name, text(value)); };
        const content = (element, value) => { if (element.textContent !== value) element.textContent = value; };
        const make = (tag, className, value) => {
            const element = document.createElement(tag);
            if (className) element.className = className;
            if (value !== undefined) element.textContent = value;
            return element;
        };
        function labelFor(select) {
            if (select.getAttribute('aria-label')) return select.getAttribute('aria-label');
            const labels = [...(select.labels || [])].map(label => {
                const copy = label.cloneNode(true);
                copy.querySelectorAll('.pm-select-wrap,select').forEach(child => child.remove());
                return copy.textContent.trim();
            }).filter(Boolean).join(' ');
            return labels || select.title || select.name || '选择选项';
        }
        function summary(select) {
            const options = [...select.options], selected = options.filter(option => option.selected);
            if (!select.multiple) return selected.length ? selected[0].textContent.trim() : '请选择…';
            const names = selected.slice(0, 2).map(option => option.textContent.trim()).join('、');
            return `已选 ${selected.length} / ${options.length}` + (names ? ' · ' + names + (selected.length > 2 ? '…' : '') : ' · 未选择');
        }
        function queue(select) {
            if (select && models.has(select)) pending.add(select);
            if (frame !== null) return;
            frame = raf(() => {
                frame = null;
                const targets = [...pending]; pending.clear();
                targets.forEach(select => syncOne(models.get(select)));
                prune();
            });
        }
        function hookProperty(model, name) {
            const select = model.select;
            if (Object.prototype.hasOwnProperty.call(select, name)) return;
            const descriptor = Object.getOwnPropertyDescriptor(win.HTMLSelectElement.prototype, name);
            if (!descriptor || !descriptor.get || !descriptor.set || !descriptor.configurable) return;
            const setter = function (value) { descriptor.set.call(this, value); queue(this); };
            try {
                Object.defineProperty(select, name, {configurable: true, enumerable: descriptor.enumerable,
                    get() { return descriptor.get.call(this); }, set: setter});
                model.hooks.push({name, setter});
            } catch (_) { /* Explicit sync remains available on restricted DOMs. */ }
        }
        function enhance(select) {
            if (models.has(select) || excluded.has(select) || select.dataset.pmSelectSkip !== undefined) return;
            const wrap = make('span', 'pm-select-wrap');
            const button = make('button', 'pm-select-trigger form-select' + (select.classList.contains('form-select-sm') ? ' form-select-sm' : ''));
            button.type = 'button'; button.id = 'pmSelectTrigger' + (++serial);
            wrap.dataset.pmSelectId = select.id || button.id;
            attr(button, 'role', 'combobox'); attr(button, 'aria-haspopup', 'listbox');
            attr(button, 'aria-expanded', 'false'); attr(button, 'aria-controls', 'pmSelectList');
            const model = {select, wrap, button, hooks: [], tabIndex: select.getAttribute('tabindex'), ariaHidden: select.getAttribute('aria-hidden')};
            models.set(select, model);
            select.before(wrap); wrap.append(select, button);
            select.classList.add('pm-select-native'); select.tabIndex = -1; attr(select, 'aria-hidden', 'true');
            hookProperty(model, 'value'); hookProperty(model, 'selectedIndex');
            syncOne(model);
        }
        function scan(root) {
            if (!root || !root.querySelectorAll) return;
            if (root.matches && root.matches('select')) enhance(root);
            root.querySelectorAll('select').forEach(enhance);
        }
        function syncOne(model) {
            if (!model) return;
            const {select, button, wrap} = model;
            const value = summary(select), disabled = isDisabled(select);
            content(button, value); button.disabled = disabled;
            wrap.hidden = select.hidden;
            attr(button, 'aria-label', labelFor(select) + '：' + value);
            attr(button, 'aria-disabled', disabled ? 'true' : 'false');
            if (current === model) {
                if (disabled || !select.isConnected || select.hidden) close(false);
                else renderOptions();
            }
        }
        function sync(root) {
            if (!root) { models.forEach(syncOne); return; }
            if (models.has(root)) { syncOne(models.get(root)); return; }
            scan(root);
            models.forEach(model => { if (root.contains && root.contains(model.select)) syncOne(model); });
        }
        function panelOnce() {
            if (panel) return;
            panel = make('div', 'pm-select-panel'); panel.id = 'pmSelectPanel'; panel.hidden = true;
            search = make('input', 'pm-select-search form-control'); search.type = 'search'; search.autocomplete = 'off';
            search.placeholder = '输入文字筛选选项'; search.dataset.pmSelectSearch = '';
            attr(search, 'aria-label', '搜索选项'); attr(search, 'role', 'combobox');
            attr(search, 'aria-controls', 'pmSelectList'); attr(search, 'aria-autocomplete', 'list');
            const actions = make('div', 'pm-select-actions');
            allButton = make('button', 'btn btn-sm btn-outline-info', '全选当前搜索结果'); allButton.type = 'button'; allButton.dataset.pmSelectAll = '';
            clearButton = make('button', 'btn btn-sm btn-outline-secondary', '清空全部选择'); clearButton.type = 'button'; clearButton.dataset.pmSelectClear = '';
            actions.append(allButton, clearButton);
            status = make('div', 'pm-select-status'); attr(status, 'role', 'status'); attr(status, 'aria-live', 'polite');
            listbox = make('div', 'pm-select-list'); listbox.id = 'pmSelectList'; listbox.dataset.pmSelectList = ''; attr(listbox, 'role', 'listbox');
            panel.append(search, actions, status, listbox); document.body.append(panel);
        }
        function position() {
            if (!current || !panel || panel.hidden) return;
            if (!current.button.isConnected || !current.button.getClientRects().length) { close(false); return; }
            const rect = current.button.getBoundingClientRect(), viewport = win.visualViewport;
            const width = viewport ? viewport.width : win.innerWidth, height = viewport ? viewport.height : win.innerHeight;
            const offsetTop = viewport ? viewport.offsetTop : 0, offsetLeft = viewport ? viewport.offsetLeft : 0;
            const panelWidth = Math.min(Math.max(rect.width, 280), Math.max(160, width - 16));
            panel.style.width = panelWidth + 'px';
            const below = offsetTop + height - rect.bottom - 8, above = rect.top - offsetTop - 8;
            const upward = below < 220 && above > below;
            const room = Math.max(100, Math.min(420, upward ? above : below, height - 16));
            panel.style.maxHeight = room + 'px';
            const actualHeight = Math.min(panel.scrollHeight, room);
            panel.style.left = Math.max(offsetLeft + 8, Math.min(rect.left, offsetLeft + width - panelWidth - 8)) + 'px';
            panel.style.top = Math.max(offsetTop + 8, upward ? rect.top - actualHeight - 4 : Math.min(rect.bottom + 4, offsetTop + height - actualHeight - 8)) + 'px';
        }
        function queuePosition(event) {
            if (!current || (event && panel && panel.contains(event.target)) || positionFrame !== null) return;
            positionFrame = raf(() => { positionFrame = null; position(); });
        }
        function highlight(option, scroll) {
            activeOption = option || null;
            if (!listbox) return;
            listbox.querySelectorAll('[data-pm-select-option]').forEach(row => {
                const active = current && current.select.options[Number(row.dataset.optionIndex)] === option;
                row.classList.toggle('pm-select-active', !!active);
                if (active) {
                    attr(search, 'aria-activedescendant', row.id);
                    if (scroll) row.scrollIntoView({block: 'nearest'});
                }
            });
            if (!option) search.removeAttribute('aria-activedescendant');
        }
        function renderOptions() {
            if (!current) return;
            const select = current.select, query = normalizeSearch(search.value), multi = select.multiple;
            const oldScroll = listbox.scrollTop, fragment = document.createDocumentFragment();
            visible = [...select.options].filter(option => !optionHidden(option) && normalizeSearch(option.textContent + ' ' + option.value + ' ' + (option.parentElement.label || '')).includes(query));
            allButton.parentElement.hidden = !multi;
            allButton.disabled = !visible.some(option => !optionDisabled(option) && !option.selected);
            clearButton.disabled = ![...select.options].some(option => option.selected && !optionDisabled(option));
            attr(listbox, 'aria-multiselectable', multi ? 'true' : 'false');
            content(status, `${visible.length} 个匹配选项` + (multi ? ` · 已选 ${[...select.selectedOptions].length} 个；搜索外的选择会保留` : ' · 方向键定位，Enter 选择'));
            visible.forEach(option => {
                const index = option.index, disabled = optionDisabled(option);
                const row = make(multi ? 'label' : 'button', 'pm-select-option');
                if (!multi) { row.type = 'button'; row.disabled = disabled; }
                row.tabIndex = -1; row.id = 'pmSelectOption' + index; row.dataset.pmSelectOption = ''; row.dataset.optionIndex = index;
                attr(row, 'role', 'option'); attr(row, 'aria-selected', option.selected ? 'true' : 'false'); attr(row, 'aria-disabled', disabled ? 'true' : 'false');
                if (multi) {
                    const check = make('input', 'pm-select-check'); check.type = 'checkbox'; check.checked = option.selected;
                    check.disabled = disabled; check.tabIndex = -1; check.dataset.pmSelectCheck = ''; attr(check, 'aria-label', option.textContent.trim());
                    row.append(check);
                }
                row.append(make('span', 'pm-select-option-label', option.textContent.trim()));
                fragment.append(row);
            });
            if (!visible.length) fragment.append(make('div', 'pm-select-empty', '没有匹配选项，请换个关键词。'));
            listbox.replaceChildren(fragment); listbox.scrollTop = oldScroll;
            const usable = visible.filter(option => !optionDisabled(option));
            highlight(usable.includes(activeOption) ? activeOption : usable.find(option => option.selected) || usable[0], false);
            position();
        }
        function open(model, seed) {
            if (!model || isDisabled(model.select) || model.select.hidden) return;
            panelOnce();
            if (current && current !== model) close(false);
            current = model;
            // Stay within the Bootstrap modal's focus trap, outside its clipped content.
            const portal = model.select.closest('.modal') || document.body;
            if (panel.parentElement !== portal) portal.append(panel);
            panel.hidden = false; search.value = seed || ''; activeOption = null;
            attr(model.button, 'aria-expanded', 'true'); attr(search, 'aria-expanded', 'true');
            attr(search, 'aria-label', '搜索' + labelFor(model.select)); attr(listbox, 'aria-label', labelFor(model.select));
            renderOptions(); search.focus({preventScroll: true});
        }
        function close(restoreFocus) {
            if (!current) return;
            const previous = current; current = null;
            attr(previous.button, 'aria-expanded', 'false');
            panel.hidden = true; attr(search, 'aria-expanded', 'false'); search.removeAttribute('aria-activedescendant');
            listbox.replaceChildren(); visible = []; activeOption = null;
            if (restoreFocus && previous.button.isConnected && !previous.button.disabled) previous.button.focus({preventScroll: true});
        }
        function notify(model, before) {
            const after = [...model.select.options].map(option => option.selected);
            if (before.length === after.length && before.every((value, index) => value === after[index])) return;
            syncOne(model);
            model.select.dispatchEvent(new win.Event('input', {bubbles: true}));
            model.select.dispatchEvent(new win.Event('change', {bubbles: true}));
        }
        function choose(option) {
            if (!current || !option || optionDisabled(option) || isDisabled(current.select)) return;
            const model = current, before = [...model.select.options].map(option => option.selected);
            if (model.select.multiple) option.selected = !option.selected;
            else { model.select.selectedIndex = option.index; close(true); }
            notify(model, before);
        }
        function chooseMany(clear) {
            if (!current || isDisabled(current.select)) return;
            const model = current, before = [...model.select.options].map(option => option.selected);
            (clear ? [...model.select.options] : visible).forEach(option => { if (!optionDisabled(option)) option.selected = !clear; });
            notify(model, before);
        }
        function modelForButton(button) { return [...models.values()].find(model => model.button === button); }
        function click(event) {
            const button = event.target.closest('.pm-select-trigger');
            if (button) { event.preventDefault(); const model = modelForButton(button); if (current === model) close(true); else open(model); return; }
            if (panel && panel.contains(event.target) && current) {
                if (event.target.closest('[data-pm-select-all]')) { event.preventDefault(); chooseMany(false); return; }
                if (event.target.closest('[data-pm-select-clear]')) { event.preventDefault(); chooseMany(true); return; }
                const row = event.target.closest('[data-pm-select-option]');
                if (row) { event.preventDefault(); choose(current.select.options[Number(row.dataset.optionIndex)]); }
                return;
            }
            const label = event.target.closest('label');
            if (label && label.control && models.has(label.control)) { event.preventDefault(); open(models.get(label.control)); }
        }
        function pointer(event) {
            if (current && !panel.contains(event.target) && !current.wrap.contains(event.target)) close(false);
        }
        function focus(event) {
            if (models.has(event.target)) { models.get(event.target).button.focus({preventScroll: true}); return; }
            if (current && !panel.contains(event.target) && !current.wrap.contains(event.target)) close(false);
        }
        function keyboard(event) {
            if (event.isComposing) return;
            const trigger = event.target.closest('.pm-select-trigger');
            if (trigger && !current && (['ArrowDown', 'ArrowUp', 'Enter', ' '].includes(event.key) || (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey))) {
                event.preventDefault(); open(modelForButton(trigger), event.key.length === 1 && event.key !== ' ' ? event.key : ''); return;
            }
            if (!current || (!panel.contains(event.target) && !current.wrap.contains(event.target))) return;
            if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(true); return; }
            if (event.key === 'Tab' && panel.contains(event.target)) {
                event.preventDefault();
                const actions = current.select.multiple ? [search, allButton, clearButton].filter(element => !element.disabled) : [search];
                const next = actions.indexOf(event.target) + (event.shiftKey ? -1 : 1);
                if (next >= 0 && next < actions.length) { actions[next].focus(); return; }
                const model = current, root = model.select.closest('.modal') || document;
                const controls = [...root.querySelectorAll('a[href],button,input,select,textarea,[tabindex]')].filter(element =>
                    element.tabIndex >= 0 && !element.matches(':disabled') && !panel.contains(element) && element.getClientRects().length);
                const destination = controls[controls.indexOf(model.button) + (event.shiftKey ? -1 : 1)] || model.button;
                close(false); destination.focus({preventScroll: true}); return;
            }
            if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                event.preventDefault();
                const options = visible.filter(option => !optionDisabled(option)), index = options.indexOf(activeOption);
                const next = options.length ? (index + (event.key === 'ArrowDown' ? 1 : -1) + options.length) % options.length : -1;
                highlight(options[next], true); search.focus({preventScroll: true});
            } else if (event.key === 'Enter' && event.target === search) { event.preventDefault(); choose(activeOption); }
        }
        function input(event) { if (event.target === search && current) { activeOption = null; renderOptions(); } }
        function change(event) { if (models.has(event.target)) syncOne(models.get(event.target)); }
        function hideModal(event) { if (current && event.target.contains(current.select)) close(false); }
        function reset(event) { win.setTimeout(() => sync(event.target), 0); }
        function release(model, permanent) {
            if (current === model) close(false);
            models.delete(model.select); pending.delete(model.select);
            model.hooks.forEach(hook => { if (Object.getOwnPropertyDescriptor(model.select, hook.name)?.set === hook.setter) delete model.select[hook.name]; });
            model.select.classList.remove('pm-select-native');
            if (model.tabIndex === null) model.select.removeAttribute('tabindex'); else model.select.setAttribute('tabindex', model.tabIndex);
            if (model.ariaHidden === null) model.select.removeAttribute('aria-hidden'); else model.select.setAttribute('aria-hidden', model.ariaHidden);
            if (model.select.parentNode === model.wrap) {
                if (model.wrap.parentNode) model.wrap.replaceWith(model.select);
                else model.button.remove();
            } else {
                // The application owns replacement/removal of the native control.
                // Remove only our proxy; never resurrect a removed old select or
                // discard a replacement which is already mounted inside the wrap.
                model.button.remove();
                if (model.wrap.parentNode) model.wrap.replaceWith(...model.wrap.childNodes);
            }
            if (permanent) excluded.add(model.select);
        }
        function prune() {
            [...models.values()].forEach(model => {
                if (!model.select.isConnected) release(model, false);
                else if (model.select.parentNode !== model.wrap) { release(model, false); enhance(model.select); }
            });
        }
        const observer = new win.MutationObserver(records => {
            let removals = false;
            records.forEach(record => {
                const element = record.target.nodeType === 1 ? record.target : record.target.parentElement;
                const select = element && element.closest('select');
                if (select && models.has(select)) queue(select);
                if (record.type === 'childList') {
                    record.addedNodes.forEach(node => { if (node.nodeType === 1) scan(node); });
                    if (record.removedNodes.length) removals = true;
                } else if (record.attributeName === 'disabled' && element && element.tagName === 'FIELDSET') element.querySelectorAll('select').forEach(queue);
            });
            if (removals) queue();
        });
        const bindings = [['click', click, false], ['pointerdown', pointer, true], ['focusin', focus, false], ['keydown', keyboard, true],
            ['input', input, false], ['change', change, false], ['hide.bs.modal', hideModal, false], ['hidden.bs.modal', hideModal, false], ['reset', reset, false], ['scroll', queuePosition, true]];
        function mount(root) {
            if (removed) return;
            if (!observing) {
                observing = true;
                bindings.forEach(([name, handler, capture]) => document.addEventListener(name, handler, capture));
                win.addEventListener('resize', queuePosition);
                if (win.visualViewport) { win.visualViewport.addEventListener('resize', queuePosition); win.visualViewport.addEventListener('scroll', queuePosition); }
                observer.observe(document.body, {childList: true, subtree: true, characterData: true, attributes: true,
                    attributeFilter: ['disabled', 'multiple', 'label', 'selected', 'hidden', 'aria-label', 'class']});
            }
            scan(root || document);
        }
        function destroy(root) {
            const targets = [...models.values()].filter(model => !root || root === model.select || (root.contains && root.contains(model.select)));
            if (!root) {
                removed = true; observer.disconnect();
                bindings.forEach(([name, handler, capture]) => document.removeEventListener(name, handler, capture));
                win.removeEventListener('resize', queuePosition);
                if (win.visualViewport) { win.visualViewport.removeEventListener('resize', queuePosition); win.visualViewport.removeEventListener('scroll', queuePosition); }
            }
            targets.forEach(model => release(model, !!root));
            if (!root && panel) panel.remove();
        }
        return {mount, sync, destroy, close: () => close(true)};
    }

    let instance = null;
    const api = {normalizeSearch, create,
        mount(root) { if (!instance) instance = create(global.document); instance.mount(root || global.document); return instance; },
        sync(root) { if (instance) instance.sync(root); },
        destroy(root) { if (instance) { instance.destroy(root); if (!root) instance = null; } }};
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    if (global.document) {
        global.ProductManagerSelects = api;
        const auto = () => { if (global.document.getElementById('pmCatalogPane') || global.document.getElementById('pmSiteSelect')) api.mount(global.document); };
        if (global.document.readyState === 'loading') global.document.addEventListener('DOMContentLoaded', auto, {once: true}); else auto();
    }
})(typeof window === 'undefined' ? globalThis : window);
