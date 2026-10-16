"""Real offline Chromium checks for the shared native-select enhancement."""
from pathlib import Path

import pytest

playwright_api = pytest.importorskip("playwright.sync_api")
ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def browser():
    with playwright_api.sync_playwright() as runtime:
        try:
            instance = runtime.chromium.launch(headless=True)
        except playwright_api.Error as exc:
            if "Executable doesn't exist" in str(exc):
                pytest.skip("Local Playwright Chromium is required; never download during tests")
            raise
        yield instance
        instance.close()


@pytest.fixture
def page(browser):
    context = browser.new_context(viewport={"width": 920, "height": 700})
    requests = []
    context.route("**/*", lambda route: (requests.append(route.request.url), route.abort()))
    current = context.new_page()
    current.set_content("""<div id='pmCatalogPane'>
        <label for='single'>库存状态</label><select id='single'><option value='a'>Éclair 苹果</option><option value='b'>草莓</option><option value='c'>柠檬</option></select>
        <label for='multi'>站点</label><select id='multi' multiple><option value='a' selected>甲站点</option><option value='b'>乙站点</option><option value='c' selected>丙站点</option></select>
        <button id='outside' style='position:fixed;bottom:8px;right:8px'>外部按钮</button><div id='dynamic'></div>
        </div>""")
    current.add_style_tag(path=str(ROOT / "static/css/product_manager_selects.css"))
    current.add_script_tag(path=str(ROOT / "static/js/product_manager_selects.js"))
    yield current
    assert requests == [], "The component must never make network requests"
    context.close()


def trigger(page, identity):
    return page.locator(f'.pm-select-wrap[data-pm-select-id="{identity}"] .pm-select-trigger')


def selected(page, identity):
    return page.locator("#" + identity).evaluate("element => [...element.selectedOptions].map(o => o.value)")


def test_native_event_identity_and_search_are_separate(page):
    page.evaluate("""() => { window.events=[]; for(const type of ['input','change'])
        document.getElementById('single').addEventListener(type,e=>events.push([e.type,e.target.id,e.target.value])); }""")
    trigger(page, "single").click()
    search = page.locator("[data-pm-select-search]")
    search.fill("草莓")
    assert selected(page, "single") == ["a"]
    assert page.evaluate("events") == []
    search.press("Enter")
    assert selected(page, "single") == ["b"]
    assert page.evaluate("events") == [["input", "single", "b"], ["change", "single", "b"]]
    assert trigger(page, "single").get_attribute("aria-expanded") == "false"


def test_empty_multiple_is_never_implicitly_all_and_preserves_hidden_selection(page):
    page.evaluate("""() => { const s=document.getElementById('multi'); for(const o of s.options)o.selected=false;
        ProductManagerSelects.sync(s); window.changes=0;s.addEventListener('change',()=>changes++); }""")
    assert "已选 0 / 3" in trigger(page, "multi").inner_text()
    trigger(page, "multi").click()
    search = page.locator("[data-pm-select-search]")
    search.fill("甲")
    assert selected(page, "multi") == []
    page.locator("[data-pm-select-all]").click()
    assert selected(page, "multi") == ["a"]
    search.fill("乙")
    page.locator("[data-pm-select-all]").click()
    assert selected(page, "multi") == ["a", "b"]
    assert page.evaluate("changes") == 2
    search.fill("不存在的站点")
    assert page.locator(".pm-select-empty").is_visible()
    assert selected(page, "multi") == ["a", "b"]
    page.locator("[data-pm-select-clear]").click()
    assert selected(page, "multi") == []
    assert page.evaluate("changes") == 3


def test_sync_keeps_query_focus_and_original_selected_values(page):
    trigger(page, "multi").click()
    search = page.locator("[data-pm-select-search]")
    search.fill("乙")
    page.evaluate("""() => { const s=document.getElementById('multi'); s.options[1].selected=true;
        ProductManagerSelects.sync(s); }""")
    assert search.input_value() == "乙"
    assert search.evaluate("e=>e===document.activeElement")
    assert selected(page, "multi") == ["a", "b", "c"]
    assert page.locator("[data-pm-select-check]").is_checked()
    assert "已选 3 / 3" in trigger(page, "multi").inner_text()


def test_optgroup_disabled_hidden_and_duplicate_values_follow_native_options(page):
    page.evaluate("""() => {document.getElementById('dynamic').innerHTML=`
        <select id="groups"><optgroup label="禁止修改" disabled><option value="duplicate">禁用项</option></optgroup>
        <optgroup label="可用"><option value="duplicate">真正可用项</option><option value="hidden" hidden>隐藏项</option></optgroup></select>`;}""")
    trigger(page, "groups").click()
    assert page.locator("[data-pm-select-option]").count() == 2
    assert page.locator('[data-option-index="0"]').is_disabled()
    page.locator("[data-pm-select-search]").fill("真正")
    page.locator("[data-pm-select-option]").click()
    assert page.locator("#groups").evaluate("s=>s.selectedIndex") == 1


def test_label_outside_click_and_disabled_fieldset(page):
    page.locator('label[for="single"]').click()
    assert page.locator("#pmSelectPanel").is_visible()
    page.locator("#outside").click()
    assert page.locator("#pmSelectPanel").is_hidden()
    page.evaluate("""() => { const f=document.createElement('fieldset'); f.id='fieldset';
        document.getElementById('dynamic').append(f);f.innerHTML='<select id="field"><option>A</option></select>'; }""")
    page.locator("#fieldset").evaluate("f=>f.disabled=true")
    page.wait_for_function("document.querySelector('[data-pm-select-id=field] button').disabled")
    assert trigger(page, "field").is_disabled()


def test_tab_leaves_single_menu_in_original_form_order(page):
    trigger(page, "single").click()
    page.locator("[data-pm-select-search]").press("Tab")
    assert page.locator("#pmSelectPanel").is_hidden()
    assert trigger(page, "multi").evaluate("e=>e===document.activeElement")


def test_shared_lazy_menu_and_observer_quiescence_for_large_table(page):
    page.evaluate("""() => {document.getElementById('dynamic').innerHTML=Array.from({length:100},(_,i)=>
        '<select id="row'+i+'"><option value="in">有货</option><option value="out">售完</option></select>').join(''); }""")
    page.wait_for_function("document.querySelectorAll('.pm-select-trigger').length===102")
    assert page.locator("#pmSelectPanel").count() == 0
    trigger(page, "row0").click()
    assert page.locator("#pmSelectPanel").count() == 1
    assert page.locator("[data-pm-select-option]").count() == 2
    trigger(page, "row99").click()
    assert page.locator("#pmSelectPanel").count() == 1
    assert page.locator("[data-pm-select-option]").count() == 2
    page.locator("[data-pm-select-search]").press("Escape")
    assert page.locator("[data-pm-select-option]").count() == 0
    page.evaluate("""() => {window.mutations=0;new MutationObserver(records=>mutations+=records.length)
        .observe(document.body,{subtree:true,childList:true,attributes:true});} """)
    page.wait_for_timeout(100)
    count = page.evaluate("mutations")
    page.wait_for_timeout(150)
    assert page.evaluate("mutations") == count, "Enhancement must settle without a MutationObserver loop"


def test_destroy_restores_native_semantics_without_duplicate_wrappers(page):
    page.evaluate("ProductManagerSelects.destroy()")
    assert page.locator(".pm-select-wrap").count() == 0
    assert page.locator("#single").is_visible()
    assert page.locator("#single").get_attribute("tabindex") is None
    assert page.locator("#single").get_attribute("aria-hidden") is None
    assert page.locator("#single").evaluate("s=>Object.getOwnPropertyDescriptor(s,'value')===undefined")
    page.evaluate("ProductManagerSelects.mount()")
    assert page.locator(".pm-select-wrap").count() == 2
    assert selected(page, "multi") == ["a", "c"]


@pytest.mark.parametrize("operation", ["remove", "replace", "move"])
def test_external_native_control_changes_never_resurrect_old_select(page, operation):
    page.evaluate("""operation => {
        const original=document.getElementById('single');window.original=original;
        if(operation==='remove') original.remove();
        if(operation==='replace') {
            const replacement=document.createElement('select'); replacement.id='replacement';
            replacement.innerHTML='<option value="new">新控件</option>'; original.replaceWith(replacement);
        }
        if(operation==='move') document.getElementById('dynamic').append(original);
    }""", operation)
    if operation == "remove":
        page.wait_for_function("document.querySelectorAll('.pm-select-wrap').length===1")
        assert page.locator("#single").count() == 0
        assert page.evaluate("!original.isConnected")
    elif operation == "replace":
        page.wait_for_function("document.querySelectorAll('[data-pm-select-id=replacement]').length===1 && document.querySelectorAll('.pm-select-wrap').length===2")
        assert page.locator("#single").count() == 0
        assert page.locator("#replacement").count() == 1
        assert "新控件" in trigger(page, "replacement").inner_text()
        assert page.evaluate("!original.isConnected")
    else:
        page.wait_for_function("document.querySelector('#dynamic .pm-select-trigger')!==null")
        assert page.locator("#single").count() == 1
        assert page.locator(".pm-select-wrap .pm-select-wrap").count() == 0
        assert selected(page, "single") == ["a"]
