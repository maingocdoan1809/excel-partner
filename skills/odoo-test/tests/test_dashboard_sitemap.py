"""Workspace sitemap APIs and their browser tab, including workspaces without test sessions."""
import json
import queue
import threading
import urllib.error
import urllib.parse
import urllib.request

import pytest

from oqa import browser, dashboard_server, sitemap, workspace
from test_sitemap_languages import BilingualRpc


class SiteMapApp:
    def __init__(self):
        self.app = dashboard_server.Dashboard(timeout=60)
        urls = queue.Queue()
        self.thread = threading.Thread(target=lambda: self.app.serve(urls.put, open_browser=False, forever=True), daemon=True)
        self.thread.start()
        self.url = urls.get(timeout=5)
        self.base = self.url.split('/?')[0]

    def get(self, path, **params):
        params.setdefault('t', self.app.token)
        request = f'{self.base}{path}?' + urllib.parse.urlencode(params)
        try:
            with urllib.request.urlopen(request, timeout=5) as response:
                return response.status, json.loads(response.read())
        except urllib.error.HTTPError as error:
            return error.code, json.loads(error.read())

    def close(self):
        self.app.quit = True
        self.thread.join(timeout=3)


@pytest.fixture
def map_app(ws):
    sitemap.save(ws['name'], sitemap.build(BilingualRpc(), profile='Admin', version='19.0'))
    app = SiteMapApp()
    yield app
    app.close()


def open_orders(page):
    page.locator('#mapDiagram').wait_for()
    page.fill('#mapQuery', 'sale.menu_orders')
    page.locator('#mapSearchForm').get_by_role('button', name='Tìm', exact=True).click()
    page.locator('.map-result[data-map-screen="menu:2"]').click()


def test_sitemap_is_available_without_a_session_and_protected_by_the_token(map_app):
    assert map_app.get('/api/sessions')[1]['sessions'] == []
    assert map_app.get('/api/workspaces')[1]['workspaces'] == [{'name': 'demo', 'language': 'vi', 'has_sitemap': True}]
    assert map_app.get('/api/sitemap', ws='demo', t='wrong')[0] == 403
    status, data = map_app.get('/api/sitemap', ws='demo', lang='ja')
    assert status == 200 and data['available'] and data['language'] == 'ja'
    assert data['meta']['menus'] == 3 and data['meta']['languages'] == ['en_US', 'ja_JP']
    menus = {menu['key']: menu for menu in data['menus']}
    assert menus['menu:2']['name'] == '注文' and menus['menu:2']['parent'] == 'menu:1'
    assert menus['menu:2']['action'] == 'window:10'
    assert 'nodes' not in data  # Do not send all fields/views to the browser just to draw the menu tree.
    assert any(node['key'] == 'window:10' and node['kind'] == 'action' for node in data['graph']['nodes'])
    assert all(isinstance(node['fields'], int) for node in data['graph']['nodes'])
    assert {'from': 'menu:2', 'to': 'window:10', 'kind': 'opens'} in data['graph']['edges']
    assert data['unresolved_count'] == 2 and '読み取れませんでした' in data['unresolved'][0]['why']


def test_search_and_screen_api_share_multilingual_cli_results(map_app):
    _, found = map_app.get('/api/sitemap', ws='demo', what='find', q='確認', kind='view', lang='en')
    assert found['results'][0]['key'] == 'view:sale.order:form:2'
    assert "button 'Confirm'" in found['results'][0]['reason']
    _, actions = map_app.get('/api/sitemap', ws='demo', what='find', kind='action')
    assert actions['results'] and all(item['kind'] == 'action' for item in actions['results'])
    _, screen = map_app.get('/api/sitemap', ws='demo', what='screen', screen='menu:2', lang='ja')
    assert screen['label'] == '販売 / 注文' and screen['views'] == ['form', 'list']
    assert 'ボタン「確認」' in screen['buttons'] and "button[name='action_confirm']" in screen['buttons']
    assert '顧客' in screen['fields'] and '開始: メニュー 販売' in screen['path']
    assert any(field['label'] == '顧客' for field in screen['controls']['fields'])
    assert [item['key'] for item in screen['route']] == ['menu:1', 'menu:2']
    _, filtered = map_app.get('/api/sitemap', ws='demo', what='screen', screen='menu:2', lang='en', view='form', match='顧客', required='true')
    assert 'Customer' in filtered['fields'] and 'Note' not in filtered['fields']


def test_workspace_isolation_empty_state_and_bad_requests(map_app, ws, store):
    workspace.save({**ws, 'name': 'other'}, store, {'Admin': 'test-admin', 'Lead': 'test-lead'})
    _, empty = map_app.get('/api/sitemap', ws='other')
    assert empty == {'workspace': 'other', 'available': False, 'language': 'en'}
    assert map_app.get('/api/workspaces')[1]['workspaces'][1]['has_sitemap'] is False
    assert map_app.get('/api/sitemap', ws='demo', what='screen', screen='menu:999')[0] == 404
    assert map_app.get('/api/sitemap', ws='demo', what='find', kind='wrong')[0] == 400
    assert map_app.get('/api/sitemap', ws='demo', what='delete')[0] == 400
    assert 'error' in map_app.get('/api/sitemap', ws='../demo')[1]
    assert 'error' in map_app.get('/api/sitemap', ws='demo', lang='fr')[1]
    sitemap.file_of('demo').write_text('{bad json', encoding='utf-8')
    assert 'error' in map_app.get('/api/sitemap', ws='demo')[1]


def test_browser_can_search_switch_language_and_open_details_without_a_session(map_app):
    pytest.importorskip('playwright')
    with browser.launch(True) as new_page:
        _, page = new_page()
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto(map_app.url)
        page.get_by_role('tab', name='site_map', exact=True).click()
        open_orders(page)
        page.locator('#mapScreenText').filter(has_text='Confirm').wait_for()
        page.select_option('#mapLanguage', 'ja')
        page.locator('#mapScreenText').filter(has_text='確認').wait_for()
        page.locator('[data-map-mode="fields"]').click()
        page.locator('#mapScreenText').filter(has_text='顧客').wait_for()
        page.check('#mapRequired')
        page.locator('#mapScreenText').filter(has_text='顧客').wait_for()
        assert '備考' not in page.inner_text('#mapScreenText')
        page.locator('[data-map-mode="path"]').click()
        page.locator('#mapScreenText .map-route-step').filter(has_text='販売').first.wait_for()
        page.fill('#mapQuery', '確認')
        page.locator('#mapSearchForm').get_by_role('button', name='Tìm', exact=True).click()
        page.locator('.map-result[data-map-screen="view:sale.order:form:2"]').click()
        page.locator('#mapScreenText').wait_for()
        assert page.get_by_role('tab', name='site_map', exact=True).get_attribute('aria-selected') == 'true'
        assert not errors, errors


@pytest.mark.parametrize('width', [375, 800, 1440])
def test_sitemap_is_responsive_and_has_named_controls(map_app, width):
    pytest.importorskip('playwright')
    with browser.launch(True) as new_page:
        _, page = new_page()
        page.set_viewport_size({'width': width, 'height': 900})
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto(map_app.url + '&tab=site_map')
        open_orders(page)
        page.locator('#mapScreenText').wait_for()
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth + 1')
        if width == 375:
            selected = page.locator('.diagram-node.selected .node-body').bounding_box()
            inspector = page.locator('#mapScreen').bounding_box()
            assert selected['y'] + selected['height'] <= inspector['y']
        assert page.evaluate('''() => ![...document.querySelectorAll('#siteMapMain button, #siteMapMain select, #siteMapMain input')].some(e => !(e.innerText || e.getAttribute('aria-label') || e.title || e.placeholder || '').trim())''')
        assert not errors, errors


def test_workspace_changes_clear_old_screens_and_empty_workspace_survives_reload(map_app, ws, store):
    pytest.importorskip('playwright')
    workspace.save({**ws, 'name': 'other'}, store, {'Admin': 'test-admin', 'Lead': 'test-lead'})
    with browser.launch(True) as new_page:
        _, page = new_page()
        page.goto(map_app.url + '&tab=site_map&workspace=demo')
        open_orders(page)
        page.locator('#mapScreenText').wait_for()
        page.select_option('#mapWorkspace', 'other')
        page.get_by_role('heading', name='Workspace chưa có sitemap').wait_for()
        assert page.locator('#mapScreenText').count() == 0
        assert 'other' in page.inner_text('#siteMapMain') and 'Customer' not in page.inner_text('#siteMapMain')
        page.reload()
        page.get_by_role('heading', name='Workspace chưa có sitemap').wait_for()
        assert page.input_value('#mapWorkspace') == 'other'
        page.select_option('#mapWorkspace', 'demo')
        page.locator('#mapDiagram').wait_for()
        assert page.locator('#mapScreenText').count() == 0


def test_progress_keeps_its_selected_case_while_sitemap_is_open(map_app, session):
    pytest.importorskip('playwright')
    with browser.launch(True) as new_page:
        _, page = new_page()
        page.goto(map_app.url)
        page.locator('tr.case[data-id="ORD-002"]').click()
        page.get_by_role('tab', name='site_map', exact=True).click()
        page.locator('#mapDiagram').wait_for()
        page.fill('#mapQuery', 'unfinished search')
        session.ask('Which customer?')
        page.wait_for_function('S.view && S.view.waiting && S.view.waiting.question === "Which customer?"')
        assert page.input_value('#mapQuery') == 'unfinished search'
        page.get_by_role('tab', name='Tiến độ test', exact=True).click()
        page.locator('.attn-panel').filter(has_text='Which customer?').wait_for()
        assert page.locator('tr.case[data-id="ORD-002"]').get_attribute('aria-selected') == 'true'


def test_late_screen_response_cannot_leak_into_another_workspace(map_app, ws, store):
    pytest.importorskip('playwright')
    workspace.save({**ws, 'name': 'other'}, store, {'Admin': 'test-admin', 'Lead': 'test-lead'})
    with browser.launch(True) as new_page:
        _, page = new_page()
        page.goto(map_app.url + '&tab=site_map&workspace=demo')
        page.locator('#mapDiagram').wait_for()
        page.evaluate('''() => {
            const original = window.fetch; window.delayedMapResponses = 0;
            window.fetch = async (...args) => {
                const response = await original(...args);
                if (String(args[0]).includes('/api/sitemap?') && String(args[0]).includes('what=screen')) {
                    await new Promise(resolve => setTimeout(resolve, 500)); window.delayedMapResponses++;
                }
                return response;
            };
        }''')
        open_orders(page)
        page.select_option('#mapWorkspace', 'other')
        page.get_by_role('heading', name='Workspace chưa có sitemap').wait_for()
        page.wait_for_function('window.delayedMapResponses > 0')
        assert 'Customer' not in page.inner_text('#siteMapMain')
        assert page.locator('#mapScreenText').count() == 0


def test_diagram_expansion_zoom_pan_and_route_highlighting(map_app):
    pytest.importorskip('playwright')
    with browser.launch(True) as new_page:
        _, page = new_page()
        page.goto(map_app.url + '&tab=site_map')
        page.locator('#mapDiagram').wait_for()
        assert page.locator('.diagram-node').count() == 2  # workspace and the Sales root
        viewport = page.locator('#mapDiagram').bounding_box()
        mini = page.locator('.diagram-minimap').bounding_box()
        assert mini['width'] <= 150 and mini['height'] <= 100  # A minimap must not cover the real graph.
        before = page.evaluate('G.zoom')
        page.locator('#mapZoomIn').click()
        assert page.evaluate('G.zoom') > before
        page.locator('#mapDiagram').focus()
        page.keyboard.press('Home')
        page.locator('[data-map-expand="menu:1"]').click()
        page.locator('.map-item[data-map-screen="menu:2"]').wait_for()
        assert page.locator('.diagram-node').count() == 3
        page.locator('.map-item[data-map-screen="menu:2"]').click()
        page.locator('#mapScreenText').filter(has_text='Confirm').wait_for()
        assert page.locator('.diagram-node.action').count() > 0 and page.locator('.diagram-node.view').count() > 0
        assert page.locator('.diagram-edge.on-route').count() > 0
        page.set_viewport_size({'width':375, 'height':900})
        page.wait_for_function('''() => {
            const node = document.querySelector('.diagram-node.selected .node-body').getBoundingClientRect();
            const panel = document.getElementById('mapScreen').getBoundingClientRect();
            return node.x >= 0 && node.right <= innerWidth && node.bottom <= panel.top;
        }''')
        page.set_viewport_size({'width':1600, 'height':900})
        page.locator('#mapCloseInspector').click()
        svg = page.locator('#mapDiagram').bounding_box()
        before = page.evaluate('[G.x, G.y]')
        page.mouse.move(svg['x'] + 25, svg['y'] + 100)
        page.mouse.down()
        page.mouse.move(svg['x'] + 125, svg['y'] + 155, steps=4)
        page.mouse.up()
        after = page.evaluate('[G.x, G.y]')
        assert after[0] == pytest.approx(before[0] + 100) and after[1] == pytest.approx(before[1] + 55)
        page.locator('[data-map-home]').first.click()
        assert page.locator('.diagram-node').count() == 2


def test_menu_with_an_unavailable_parent_is_grouped_without_inventing_a_route(map_app):
    data = sitemap.load('demo')
    data['nodes']['menu:99'] = {'kind': 'menu', 'label': 'Detached menu', 'name': 'Detached menu'}
    data['edges'].append({'from': 'menu:999', 'to': 'menu:99', 'kind': 'menu'})
    data['meta']['menus'] += 1
    sitemap.save('demo', data)
    _, screen = map_app.get('/api/sitemap', ws='demo', what='screen', screen='menu:99')
    assert screen['reachable'] is False and screen['route'] == []
    pytest.importorskip('playwright')
    with browser.launch(True) as new_page:
        _, page = new_page()
        page.goto(map_app.url + '&tab=site_map')
        page.locator('[data-map-focus="other-menus"]').click()
        page.locator('.map-item[data-map-screen="menu:99"]').click()
        page.locator('#mapScreenText').wait_for()
        page.locator('[data-map-mode="path"]').click()
        assert 'no way found' in page.inner_text('#mapScreenText')
