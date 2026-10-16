"""Independent multiselect/search acceptance; synthetic API and browser only."""
from pathlib import Path
import html
import shutil
import subprocess

import pytest

# Reuse the previous release's synthetic DB fixture, never production app state.
from test_catalog_load_scope_contract import scope_system, metadata, read_page


ROOT = Path(__file__).resolve().parents[1]


def test_non_adjacent_ids_stay_a_set_and_intersect_manager_country(scope_system):
    chosen = "1,3,7"
    data = metadata(scope_system, site_ids=chosen).get_json()
    assert {site["id"] for site in data["sites"]} == {1, 3, 7}
    data = metadata(scope_system, site_ids=chosen, manager="Alice", country="PL").get_json()
    assert [site["id"] for site in data["sites"]] == [1]
    response = read_page(scope_system, site_id=1, site_ids=chosen, manager="Alice", country="PL")
    assert response.get_json()["complete_page"] is True
    response = read_page(scope_system, site_id=3, site_ids=chosen, manager="Alice", country="PL")
    assert response.status_code == 403 and response.get_json()["code"] == "scope_mismatch"
    assert [call["site_id"] for call in scope_system.calls] == [1]


def test_empty_checkbox_csv_is_none_and_never_every_site(scope_system):
    assert metadata(scope_system, site_ids="").get_json()["sites"] == []
    assert read_page(scope_system, site_ids="").status_code == 403
    assert scope_system.calls == []


def test_checkbox_ids_cannot_restore_revoked_or_archived_sites(scope_system):
    scope_system.user.own_scoped = True
    assert {site["id"] for site in metadata(scope_system, site_ids="1,3,6").get_json()["sites"]} == {1}
    with scope_system.connect() as conn:
        conn.execute("UPDATE sites SET manager='Bob' WHERE id=1")
    assert metadata(scope_system, site_ids="1,3,6").get_json()["sites"] == []
    assert read_page(scope_system, site_id=1, site_ids="1,3,6").status_code == 403
    assert read_page(scope_system, site_id=6, site_ids="1,3,6").status_code == 409
    assert scope_system.calls == []


@pytest.mark.skipif(shutil.which("node") is None, reason="Node.js is required")
def test_real_controller_multiselect_editor_release_contracts():
    completed = subprocess.run(
        [shutil.which("node"), str(ROOT / "tests/catalog_multiselect_harness.cjs")],
        cwd=ROOT, capture_output=True, text=True, encoding="utf-8", timeout=30,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert "Independent multiselect contracts passed" in completed.stdout


@pytest.fixture(scope="module")
def local_chromium():
    playwright = pytest.importorskip("playwright.sync_api")
    with playwright.sync_playwright() as runtime:
        try:
            browser = runtime.chromium.launch(headless=True, args=["--disable-background-networking"])
        except playwright.Error as exc:
            if "Executable doesn't exist" in str(exc):
                pytest.skip("Local Playwright Chromium is required; never download during tests")
            raise
        yield browser
        browser.close()


WIDGET_HTML = """<!doctype html><html><head><meta charset="utf-8"></head><body>
<main id="pmCatalogPane"><label for="pmSiteSelect">单站站点</label>
<select id="pmSiteSelect" name="site_id"><option value="a" selected>Alpha site</option>
<option value="b">Ízesítés Łódź Blue-Ice</option><option value="disabled" disabled>Disabled item</option>
<option value="evil">EVIL_LABEL</option></select>
<label for="multi">多选站点</label><select id="multi" multiple>
<option value="1" selected>West one</option><option value="3">East three</option>
<option value="7" selected>West seven</option><option value="9">East nine</option></select>
<label for="disabled">禁用状态</label><select id="disabled" disabled><option value="no" selected>No change</option><option value="yes">Yes</option></select>
<div id="dynamic-host"></div></main></body></html>""".replace(
    "EVIL_LABEL", html.escape('<img src="https://evil.invalid/x" onerror="window.__xss=1">')
)


@pytest.fixture
def widget_page(local_chromium):
    context = local_chromium.new_context(viewport={"width": 1100, "height": 800})
    blocked = []

    def deny_network(route):
        blocked.append(route.request.url)
        route.abort()

    context.route("**/*", deny_network)
    page = context.new_page()
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    page.set_content(WIDGET_HTML)
    page.evaluate("""() => {
        window.__events = []; window.__fetches = [];
        document.addEventListener('change', e => { if (e.target.matches('select')) window.__events.push({id:e.target.id,value:e.target.value}); });
        window.fetch = (...args) => { window.__fetches.push(String(args[0])); throw Error('No network allowed'); };
    }""")
    page.add_style_tag(path=str(ROOT / "static/css/product_manager_selects.css"))
    page.add_script_tag(path=str(ROOT / "static/js/product_manager_selects.js"))
    page.evaluate("ProductManagerSelects.mount(document)")
    page.locator('.pm-select-wrap[data-pm-select-id="pmSiteSelect"] button.pm-select-trigger').wait_for()
    yield page
    assert errors == [], errors
    assert blocked == [], blocked
    assert page.evaluate("window.__fetches") == []
    context.close()


def trigger(page, identity):
    return page.locator(f'.pm-select-wrap[data-pm-select-id="{identity}"] button.pm-select-trigger')


def search_input(page):
    return page.locator("#pmSelectPanel [data-pm-select-search]")


def selected_values(page, identity):
    return page.locator("#" + identity).evaluate("element => [...element.selectedOptions].map(option => option.value)")


def test_widget_search_unicode_keyboard_and_xss_do_not_mutate_before_commit(widget_page):
    page = widget_page
    trigger(page, "pmSiteSelect").click()
    search_input(page).fill("lodz")
    page.locator('#pmSelectPanel [data-pm-select-option]').filter(has_text="Łódź").wait_for(state="visible")
    assert selected_values(page, "pmSiteSelect") == ["a"]
    assert page.evaluate("window.__events.length") == 0
    search_input(page).press("ArrowDown")
    search_input(page).press("Enter")
    assert selected_values(page, "pmSiteSelect") == ["b"]
    assert page.evaluate("window.__events.length") == 1
    trigger(page, "pmSiteSelect").click()
    search_input(page).fill("img")
    assert page.locator("#pmSelectPanel img, #pmSelectPanel svg").count() == 0
    assert page.evaluate("window.__xss || 0") == 0
    search_input(page).press("Escape")
    assert page.evaluate("document.activeElement.classList.contains('pm-select-trigger')")


def test_widget_multiple_keeps_hidden_checks_and_clear_means_empty(widget_page):
    page = widget_page
    trigger(page, "multi").click()
    search_input(page).fill("East")
    assert selected_values(page, "multi") == ["1", "7"]
    assert page.evaluate("window.__events.length") == 0
    page.locator('#pmSelectPanel [data-pm-select-option]').filter(has_text="East three").locator('input[type="checkbox"]').check()
    assert selected_values(page, "multi") == ["1", "3", "7"]
    page.locator("#pmSelectPanel [data-pm-select-all]").click()
    assert selected_values(page, "multi") == ["1", "3", "7", "9"]
    page.locator("#pmSelectPanel [data-pm-select-clear]").click()
    assert selected_values(page, "multi") == []
    assert page.evaluate("window.__events.filter(event => event.id === 'multi').length") == 3


def test_widget_programmatic_disabled_and_dynamic_options_sync_without_fake_change(widget_page):
    page = widget_page
    assert trigger(page, "disabled").is_disabled()
    page.evaluate("document.getElementById('pmSiteSelect').value = 'b'")
    page.wait_for_function("document.querySelector('[data-pm-select-id=pmSiteSelect] .pm-select-trigger').textContent.includes('Łódź')")
    page.evaluate("document.getElementById('pmSiteSelect').selectedIndex = 0")
    page.wait_for_function("document.querySelector('[data-pm-select-id=pmSiteSelect] .pm-select-trigger').textContent.includes('Alpha')")
    page.evaluate("""() => {
        const element = document.getElementById('pmSiteSelect');
        element.replaceChildren(new Option('New Mango', 'new', true, true), new Option('New Lemon', 'lemon'));
        document.getElementById('disabled').disabled = false;
    }""")
    page.wait_for_function("document.querySelector('[data-pm-select-id=pmSiteSelect] .pm-select-trigger').textContent.includes('Mango')")
    assert page.locator('[data-pm-select-id="pmSiteSelect"]').count() == 1
    assert trigger(page, "disabled").is_enabled()
    assert page.evaluate("window.__events.length") == 0
    trigger(page, "pmSiteSelect").click()
    search_input(page).fill("Lemon")
    page.locator('#pmSelectPanel [data-pm-select-option]').filter(has_text="New Lemon").click()
    assert selected_values(page, "pmSiteSelect") == ["lemon"]
    assert page.evaluate("window.__events.length") == 1


def test_dynamic_modal_select_search_retains_focus_inside_modal(widget_page):
    page = widget_page
    page.evaluate("""() => {
        document.getElementById('dynamic-host').innerHTML = '<div class="modal show" id="test-modal" style="display:block;padding:20px" tabindex="-1"><label for="modal-state">动态库存状态</label><select id="modal-state" name="stock_status"><option value="instock">有货</option><option value="outofstock">售完</option></select></div>';
    }""")
    trigger(page, "modal-state").wait_for()
    trigger(page, "modal-state").click()
    assert page.locator("#test-modal #pmSelectPanel").count() == 1
    search_input(page).fill("售完")
    assert page.evaluate("document.getElementById('test-modal').contains(document.activeElement)")
    page.locator('#pmSelectPanel [data-pm-select-option]').filter(has_text="售完").click()
    assert selected_values(page, "modal-state") == ["outofstock"]
    assert page.evaluate("window.__events.at(-1).id") == "modal-state"


def test_real_checkbox_widget_drives_catalog_without_search_side_effects(widget_page):
    page = widget_page
    filters = "".join(
        f'<label for="pmCatalog{name}Filter">{name}</label><select id="pmCatalog{name}Filter"><option value="">全部</option></select>'
        for name in ("Manager", "Country", "Site", "Brand", "Product", "Flavor", "Stock", "Publish")
    )
    markup = f"""<main id="pmCatalogPane">
        <label for="pmCatalogLoadManager">负责人</label><select id="pmCatalogLoadManager"></select>
        <label for="pmCatalogLoadCountry">国家</label><select id="pmCatalogLoadCountry"></select>
        <label for="pmCatalogLoadSite">多选站点</label><select id="pmCatalogLoadSite" multiple></select>
        <span id="pmCatalogScopeCount"></span><span id="pmCatalogScopeStatus"></span>
        <span id="pmCatalogLoadedScope"></span><button id="pmCatalogReloadSites">刷新站点</button>
        <select id="pmCatalogQueryMode"><option value="all" selected>全部产品</option><option value="brand">品牌</option></select>
        <div id="pmCatalogBrandInputWrap"><input id="pmCatalogBrandInput"></div>
        <div id="pmCatalogFlavorInputWrap"><input id="pmCatalogSearch"></div>
        <button id="pmCatalogLoad">加载</button><button id="pmCatalogStop">停止</button>
        <button id="pmCatalogRetry">重试</button><button id="pmCatalogRefresh">重读</button>
        <div id="pmCatalogStatus"></div><div id="pmCatalogCounts"></div><div id="pmCatalogProgress"></div>
        {filters}<input id="pmCatalogTextFilter"><button id="pmCatalogResetFilters">重置</button>
        <div id="pmCatalogResults"></div></main>"""
    page.evaluate("markup => document.body.innerHTML = markup", markup)
    page.evaluate("""() => {
        window.__catalogCalls = [];
        const sites = [
            {id:1,url:'https://one.invalid',manager:'Alice',country:'PL'},
            {id:3,url:'https://three.invalid',manager:'Alice',country:'PL'},
            {id:7,url:'https://seven.invalid',manager:'Alice',country:'CZ'},
            {id:9,url:'https://nine.invalid',manager:'Bob',country:'PL'}];
        const reply = data => ({ok:true,status:200,headers:{get:()=>'application/json'},json:async()=>data});
        window.fetch = async (url, options={}) => {
            const request = new URL(url,'https://local.invalid'), p=request.searchParams;
            window.__catalogCalls.push({path:request.pathname,site:p.get('site_id'),method:options.method||'GET'});
            if (request.pathname.endsWith('/catalog-sites')) {
                const ids=p.has('site_ids')?p.get('site_ids').split(','):null;
                return reply({sites:sites.filter(site=>(!ids||ids.includes(String(site.id)))&&(!p.get('manager')||site.manager===p.get('manager'))&&(!p.get('country')||site.country===p.get('country')))});
            }
            if (request.pathname.endsWith('/catalog-page')) {
                const id=Number(p.get('site_id'));
                return reply({site_id:id,parent_id:0,page:1,per_page:100,source_ids:[10],source_statuses:{'10':'publish'},variable_products:[],total:1,total_pages:1,scanned:1,complete_page:true,has_more:false,next_page:null,warnings:[],rows:[{
                    site_id:id,product_id:10,variation_id:0,type:'simple',name:'Mint',product_name:'Merrymi 12000',brands:['Merrymi'],flavors:['Mint'],sku:'synthetic-'+id,status:'publish',manage_stock:true,stock_quantity:5,stock_status:'instock',price:'10',regular_price:'10',sale_price:'',
                    edit_identity:{sku:'synthetic-'+id,type:'simple',parent_id:0,attributes:[]},edit_before:{manage_stock:true,stock_quantity:5,stock_status:'instock',regular_price:'10',sale_price:''}}]});
            }
            if (request.pathname.endsWith('/catalog-edit-config')) return reply({csrf_token:'synthetic-only',sites:sites.map(site=>({...site,can_edit:true,can_clone:true})),clone_targets:[]});
            throw Error('Unexpected catalog side effect '+url);
        };
    }""")
    page.add_script_tag(path=str(ROOT / "static/js/product_manager_catalog.js"))
    page.wait_for_function("document.getElementById('pmCatalogPane').productManagerCatalog?.getSnapshot().metadataPhase === 'ready'")
    assert selected_values(page, "pmCatalogLoadSite") == ["1", "3", "7", "9"]
    assert page.evaluate("window.__catalogCalls.length") == 1
    trigger(page, "pmCatalogLoadSite").click()
    page.locator("#pmSelectPanel [data-pm-select-clear]").click()
    assert selected_values(page, "pmCatalogLoadSite") == []
    page.wait_for_function("document.getElementById('pmCatalogLoad').disabled")
    search_input(page).fill("one.invalid")
    page.locator('#pmSelectPanel [data-pm-select-option]').filter(has_text="one.invalid").locator('input[type="checkbox"]').check()
    search_input(page).fill("seven.invalid")
    page.locator('#pmSelectPanel [data-pm-select-option]').filter(has_text="seven.invalid").locator('input[type="checkbox"]').check()
    search_input(page).press("Escape")
    assert selected_values(page, "pmCatalogLoadSite") == ["1", "7"]
    assert page.evaluate("window.__catalogCalls.length") == 1
    page.locator("#pmCatalogLoad").click()
    page.wait_for_function("document.getElementById('pmCatalogPane').productManagerCatalog.getSnapshot().phase === 'complete'")
    assert page.evaluate("window.__catalogCalls.filter(call=>call.path.endsWith('/catalog-page')).map(call=>call.site).sort()") == ["1", "7"]
    page.add_script_tag(path=str(ROOT / "static/js/product_manager_catalog_editor.js"))
    page.evaluate("""async () => {
        const catalog=document.getElementById('pmCatalogPane').productManagerCatalog;
        window.__editor=ProductManagerCatalogEditor.createEditor({catalog,fetch:window.fetch,autoPoll:false,confirm:()=>false});
        await window.__editor.ready(); window.__editor.select(['1:10:0','7:10:0']); window.__editor.setDraft('1:10:0','regular_price','14.5');
        window.__loadedBefore={scope:catalog.getSnapshot().loadedScope,at:catalog.getSnapshot().loadedAt};
    }""")
    before = page.evaluate("window.__catalogCalls.length")
    trigger(page, "pmCatalogCountryFilter").click()
    search_input(page).fill("does not match any option")
    assert page.evaluate("window.__editor.getSnapshot().selectedCount") == 2
    assert page.evaluate("window.__editor.getDraft('1:10:0').regular_price") == "14.5"
    search_input(page).press("Escape")
    trigger(page, "pmCatalogLoadSite").click()
    search_input(page).fill("nine.invalid")
    assert page.evaluate("window.__catalogCalls.length") == before
    assert selected_values(page, "pmCatalogLoadSite") == ["1", "7"]
    page.locator('#pmSelectPanel [data-pm-select-option]').filter(has_text="nine.invalid").locator('input[type="checkbox"]').check()
    search_input(page).press("Escape")
    assert page.evaluate("window.__catalogCalls.length") == before
    assert page.evaluate("""() => {
        const snapshot=document.getElementById('pmCatalogPane').productManagerCatalog.getSnapshot();
        return JSON.stringify(snapshot.loadedScope)===JSON.stringify(window.__loadedBefore.scope)&&snapshot.loadedAt===window.__loadedBefore.at;
    }""")
    assert page.evaluate("window.__editor.getSnapshot().selectedCount") == 2
    assert page.evaluate("window.__editor.getSnapshot().draftCount") == 1
