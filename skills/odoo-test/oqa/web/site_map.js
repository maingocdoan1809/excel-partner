/* Workspace graph viewer. SVG coordinates stay separate from pan/zoom and inspector state. */
const G = { focus: "", expanded: new Set(), route: [], inspector: false, searchOpen: false, x: 0, y: 0, zoom: 1, fit: true, layout: null, drag: null };
const graphNodeWidth = 280, graphNodeHeight = 96;

function graphIndex() {
  const graph = M.data && M.data.graph || { nodes: [], edges: [] };
  const nodes = new Map(graph.nodes.map(node => [node.key, node])), outgoing = new Map(), incoming = new Map();
  const orphans = [];
  graph.edges.forEach(edge => {
    if (edge.kind === "menu" && nodes.has(edge.to) && !nodes.has(edge.from)) orphans.push(edge.to);
    if (!nodes.has(edge.from) || !nodes.has(edge.to)) return;
    if (!outgoing.has(edge.from)) outgoing.set(edge.from, []);
    if (!incoming.has(edge.to)) incoming.set(edge.to, []);
    outgoing.get(edge.from).push(edge); incoming.get(edge.to).push(edge);
  });
  const edges = [...graph.edges];
  if (orphans.length) {
    const key = "other-menus";
    nodes.set(key, { key, kind: "group", label: M.lang === "ja" ? "その他のメニュー" : "Other menus", model: M.lang === "ja" ? `${orphans.length} メニュー · 親メニュー未取得` : `${orphans.length} menus · parent unavailable` });
    outgoing.set(key, orphans.map(to => ({ from: key, to, kind: "menu" })));
    outgoing.get(key).forEach(edge => { incoming.set(edge.to, [edge]); edges.push(edge); });
  }
  return { nodes, outgoing, incoming, edges };
}

function graphAncestors(key, index) {
  const queue = [[key]], seen = new Set([key]);
  for (let cursor = 0; cursor < queue.length; cursor++) {
    const chain = queue[cursor], current = chain[0], node = index.nodes.get(current);
    const incoming = (index.incoming.get(current) || []).filter(edge => ["menu", "opens", "shows", "button"].includes(edge.kind));
    if (node && ["menu", "group"].includes(node.kind) && !incoming.some(edge => edge.kind === "menu")) return chain;
    for (const edge of incoming) if (!seen.has(edge.from)) { seen.add(edge.from); queue.push([edge.from, ...chain]); }
  }
  return [key];
}

function revealGraph(key) {
  const index = graphIndex(), chain = graphAncestors(key, index);
  if (!chain.includes(G.focus)) G.focus = chain[0];
  chain.forEach(item => G.expanded.add(item));
  const node = index.nodes.get(key);
  if (node && node.kind === "menu") {
    (index.outgoing.get(key) || []).filter(edge => edge.kind === "opens").forEach(edge => G.expanded.add(edge.to));
  }
  G.fit = true;
}

function graphLayout() {
  const index = graphIndex(), positions = new Map(), links = [];
  if (!G.focus || !index.nodes.has(G.focus)) {
    const roots = [...index.nodes.values()].filter(node => ["menu", "group"].includes(node.kind) && !(index.incoming.get(node.key) || []).some(edge => edge.kind === "menu"));
    const columns = roots.length > 6 ? 2 : 1, rows = Math.max(1, Math.ceil(roots.length / columns));
    positions.set("workspace", { key: "workspace", kind: "workspace", label: M.ws, x: 0, y: (rows - 1) * 136 / 2 });
    roots.forEach((node, i) => {
      positions.set(node.key, { ...node, x: 420 + Math.floor(i / rows) * 380, y: (i % rows) * 136 });
      links.push({ from: "workspace", to: node.key, kind: "menu", overview: true });
    });
  } else {
    const visited = new Set([G.focus]), children = new Map(), tree = new Set();
    function visit(key) {
      const next = [];
      if (G.expanded.has(key)) for (const edge of index.outgoing.get(key) || []) {
        if (!visited.has(edge.to)) { visited.add(edge.to); next.push(edge.to); tree.add(edge); visit(edge.to); }
      }
      children.set(key, next);
    }
    visit(G.focus);
    const heights = new Map();
    function measure(key) {
      const height = Math.max(132, (children.get(key) || []).reduce((sum, child) => sum + measure(child), 0));
      heights.set(key, height); return height;
    }
    function place(key, depth, top) {
      positions.set(key, { ...index.nodes.get(key), x: depth * 380, y: top + (heights.get(key) - graphNodeHeight) / 2 });
      for (const child of children.get(key) || []) { place(child, depth + 1, top); top += heights.get(child); }
    }
    measure(G.focus); place(G.focus, 0, 0);
    index.edges.forEach(edge => { if (positions.has(edge.from) && positions.has(edge.to)) links.push({ ...edge, secondary: !tree.has(edge) }); });
  }
  const values = [...positions.values()];
  const bounds = { x: -40, y: -80, width: Math.max(600, ...values.map(node => node.x + graphNodeWidth)) + 80,
    height: Math.max(300, ...values.map(node => node.y + graphNodeHeight)) + 135 };
  return { positions, links, bounds, index };
}

function svgLines(label, limit = 25) {
  const chars = [...String(label)], lines = [""], width = char => /[^\u0000-\u02ff]/.test(char) ? 1.7 : 1;
  let length = 0;
  chars.forEach(char => { if (length + width(char) > limit && lines.length < 2) { lines.push(""); length = 0; } lines[lines.length - 1] += char; length += width(char); });
  if (length > limit) { let n = 0; lines[1] = [...lines[1]].filter(char => (n += width(char)) < limit - 1).join("") + "…"; }
  return lines;
}

function graphSvg(layout) {
  const routeEdges = new Set(G.route.map(edge => `${edge.from}|${edge.to}|${edge.kind}`));
  const paths = layout.links.map(edge => {
    const a = layout.positions.get(edge.from), b = layout.positions.get(edge.to);
    const x = a.x + graphNodeWidth, y = a.y + 48, tx = b.x, ty = b.y + 48;
    let path;
    if (edge.overview) path = b.x > 700 ? `M${x},${y}H340V-35H760V${ty}H${tx}` : `M${x},${y}H340V${ty}H${tx}`;
    else if (tx > x) path = `M${x},${y}C${x + (tx - x)/2},${y} ${x + (tx - x)/2},${ty} ${tx},${ty}`;
    else { const end = b.x + graphNodeWidth, rail = Math.max(x, end) + 65; path = `M${x},${y}C${rail},${y} ${rail},${ty} ${end},${ty}`; }
    const active = routeEdges.has(`${edge.from}|${edge.to}|${edge.kind}`);
    const label = !edge.overview && !edge.secondary ? ({ menu: "menu", opens: "opens", shows: "shows", button: edge.label || "button", row: "open row" }[edge.kind] || "") : "";
    return `<g><path class="diagram-edge${edge.secondary ? " secondary" : ""}${active ? " on-route" : ""}" d="${path}" marker-end="url(#${active ? "map-arrow-active" : "map-arrow"})"/>${label ? `<text class="diagram-edge-label" x="${(x + tx)/2}" y="${ty - 10}" text-anchor="middle">${esc(label.length > 26 ? label.slice(0,25) + "…" : label)}</text>` : ""}</g>`;
  }).join("");
  const nodes = [...layout.positions.values()].map(node => {
    const outgoing = layout.index.outgoing.get(node.key) || [];
    const expandable = outgoing.length > 0 && node.kind !== "workspace";
    const subtitle = node.kind === "workspace" ? `${M.data.meta.menus} menus · ${M.data.meta.actions} actions · ${M.data.meta.views} views`
      : node.kind === "view" ? `${node.fields} fields · ${node.buttons} buttons` : node.model || node.xmlid || `${outgoing.length} connections`;
    const lines = svgLines(node.label);
    return `<g class="diagram-node ${node.kind}${node.key === M.key ? " selected" : ""}" transform="translate(${node.x},${node.y})">
      <g class="map-item" ${node.kind === "workspace" ? 'data-map-home="true"' : node.kind === "group" ? `data-map-focus="${esc(node.key)}"` : `data-map-screen="${esc(node.key)}"`} role="button" tabindex="0" aria-label="${esc(node.trail || node.label)}" aria-current="${node.key === M.key}">
        <rect class="node-body" width="280" height="96" rx="11"/><rect class="node-strip" x="0" y="19" width="3" height="58" rx="1.5"/>
        <text class="node-type" x="20" y="22">${node.kind === "workspace" ? "WORKSPACE" : esc(node.kind.toUpperCase()) + (node.kind === "view" ? " / " + esc(node.type) : "")}</text>
        <text class="node-title" x="20" y="${lines.length === 1 ? 48 : 44}">${lines.map((line, i) => `<tspan x="20" dy="${i ? 18 : 0}">${esc(line)}</tspan>`).join("")}</text>
        <text class="node-sub" x="20" y="81">${esc(subtitle.length > 38 ? subtitle.slice(0,37) + "…" : subtitle)}</text><title>${esc(node.trail || node.label)}${node.xmlid ? " · " + esc(node.xmlid) : ""}</title>
      </g>${expandable ? `<g class="node-expander" data-map-expand="${esc(node.key)}" role="button" tabindex="0" aria-label="${G.expanded.has(node.key) && G.focus ? "Thu gọn" : "Mở nhánh"} ${esc(node.label)}"><circle cx="280" cy="48" r="12"/><text x="280" y="54">${G.expanded.has(node.key) && G.focus ? "−" : "+"}</text></g>` : ""}</g>`;
  }).join("");
  return `<defs><pattern id="map-dots" width="22" height="22" patternUnits="userSpaceOnUse"><circle cx="1" cy="1" r="1" fill="var(--map-grid)"/></pattern><marker id="map-arrow" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto"><path d="M0 0L7 3.5L0 7Z" fill="var(--map-edge)"/></marker><marker id="map-arrow-active" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto"><path d="M0 0L7 3.5L0 7Z" fill="var(--map-blue)"/></marker></defs>
    <rect width="100%" height="100%" fill="url(#map-dots)"/><g id="mapGraphViewport">${paths}${nodes}</g>`;
}

function graphTransform() {
  const viewport = $("mapGraphViewport"), svg = $("mapDiagram");
  if (!viewport || !svg) return;
  viewport.setAttribute("transform", `translate(${G.x},${G.y}) scale(${G.zoom})`);
  if ($("mapZoomLevel")) $("mapZoomLevel").textContent = Math.round(G.zoom * 100) + "%";
  const mini = $("mapMiniViewport");
  if (mini) { mini.setAttribute("x", -G.x/G.zoom); mini.setAttribute("y", -G.y/G.zoom); mini.setAttribute("width", svg.clientWidth/G.zoom); mini.setAttribute("height", svg.clientHeight/G.zoom); }
}

function graphFit(initial = false) {
  const svg = $("mapDiagram");
  if (!svg || !G.layout || !svg.clientWidth) return;
  const b = G.layout.bounds;
  let width = svg.clientWidth, height = svg.clientHeight, top = 65;
  const inspector = $("mapScreen");
  if (G.inspector && inspector && window.innerWidth <= 767) { height = Math.max(180, inspector.offsetTop - 15); top = 135; }
  else if (G.inspector && inspector && window.innerWidth <= 1100) width = Math.max(210, width - inspector.offsetWidth - 30);
  G.zoom = Math.min(1.15, (width - 80)/b.width, (height - top - 45)/b.height);
  const minimum = initial ? (svg.clientWidth < 600 ? .58 : .65) : .15;
  const clamped = G.zoom < minimum;
  G.zoom = Math.max(minimum, G.zoom);
  G.x = (width - b.width * G.zoom)/2 - b.x*G.zoom;
  G.y = (height - b.height * G.zoom)/2 - b.y*G.zoom + top/2;
  const selected = G.layout.positions.get(M.key);
  if (clamped && selected) {
    G.x = width/2 - (selected.x + graphNodeWidth/2)*G.zoom;
    G.y = (top + height)/2 - (selected.y + graphNodeHeight/2)*G.zoom;
  }
  graphTransform();
}

function graphZoom(factor, x, y) {
  const svg = $("mapDiagram");
  if (!svg) return;
  x = x ?? svg.clientWidth/2; y = y ?? svg.clientHeight/2;
  const next = Math.max(.15, Math.min(2.5, G.zoom * factor)), ratio = next/G.zoom;
  G.x = x - (x - G.x)*ratio; G.y = y - (y - G.y)*ratio; G.zoom = next; graphTransform();
}

function resetGraph() { G.focus = ""; G.expanded.clear(); G.inspector = false; G.route = []; G.fit = true; M.key = ""; M.screen = null; renderMap(); }

function selectGraphScreen(key) {
  if (M.key !== key) { M.view = ""; M.match = ""; M.required = false; }
  M.key = key; M.screen = null; G.inspector = true; G.searchOpen = false; revealGraph(key); loadMapScreen();
}

function renderMapScreen() {
  const screen = M.screen, japanese = M.lang === "ja";
  const title = screen ? screen.label : (graphIndex().nodes.get(M.key) || {}).label || "Chi tiết màn hình";
  let content = '<p class="map-note" role="status">Đang tải màn hình…</p>';
  if (M.detailError) content = `<div class="map-error" role="alert">${esc(M.detailError)}</div>`;
  else if (screen && !M.detailLoading) {
    const controls = screen.controls || { buttons: [], fields: [] };
    if (M.mode === "path") content = (screen.route || []).map((step, i) => `<div class="map-route-step"><span class="number">${i + 1}</span><div><small>${esc(step.via === "start" ? (japanese ? "開始" : "Start") : step.via)}</small><h3>${esc(step.label)}</h3><small>${esc(step.key)}</small></div></div>`).join("") || `<p class="map-note">${esc(screen.path)}</p>`;
    else content = (controls[M.mode] || []).map(item => {
      const flags = [];
      if (item.required) flags.push(item.required === true ? (japanese ? "必須" : "Required") : `${japanese ? "必須条件" : "Required if"}: ${item.required}`);
      if (item.readonly) flags.push(japanese ? "読み取り専用" : "Read only");
      return `<article class="map-control"><h3>${esc(item.label || item.name)}</h3><div class="technical">${esc(item.name)}</div><div class="map-badges"><span class="map-badge type">${esc(item.type)}</span><span class="map-badge">${esc(item.view)}</span>${flags.map(flag => `<span class="map-badge">${esc(flag)}</span>`).join("")}</div>
        ${item.target_label ? `<p>→ ${esc(item.target_label)}</p>` : item.type === "object" ? `<p>${japanese ? "メソッド呼び出し · 遷移先は実行時に判明" : "Method call · destination known when run"}</p>` : ""}
        ${item.relation ? `<p>→ ${esc(item.relation)}</p>` : ""}${item.invisible ? `<p>${japanese ? "表示条件" : "Shown if"}: not (${esc(item.invisible)})</p>` : ""}
        ${item.options ? `<p>${item.options.map(option => esc(option)).join(" · ")}</p>` : ""}<code>${esc(item.selector)}</code></article>`;
    }).join("");
    if (!content) content = `<p class="map-note">${M.match ? "Không có mục khớp bộ lọc." : "Không có mục nào được khai báo trong sitemap của màn hình này."}</p>`;
    if (controls[M.mode + "_total"] > 60) content += '<p class="map-note">Hiển thị 60 mục đầu. Dùng bộ lọc để thu hẹp.</p>';
  }
  return `<div class="map-inspector-head"><div><div class="map-eyebrow">SCREEN INSPECTOR</div><h2>${esc(title)}</h2><div class="map-note">${esc(M.key)}</div></div><button class="map-close" id="mapCloseInspector" aria-label="Đóng chi tiết">×</button></div>
    <div class="map-inspector-body"><div class="tabs" role="tablist" aria-label="Chi tiết sitemap">${[["buttons", "Nút"], ["fields", "Trường"], ["path", "Đường đi"]].map(([key, label]) => `<button role="tab" data-map-mode="${key}" aria-selected="${M.mode === key}">${label}</button>`).join("")}</div>
    ${screen && M.mode !== "path" ? `<form id="mapDetailForm" class="map-toolbar"><label class="map-query"><input id="mapMatch" aria-label="Lọc nút hoặc trường" value="${esc(M.match)}" placeholder="Lọc Anh / Nhật…"></label><label><select id="mapView" aria-label="Loại view"><option value="">Tự chọn</option>${screen.views.map(type => `<option ${M.view === type ? "selected" : ""}>${esc(type)}</option>`).join("")}</select></label>${M.mode === "fields" ? `<label class="map-required"><input type="checkbox" id="mapRequired" ${M.required ? "checked" : ""}>Bắt buộc</label>` : ""}<button class="btn" type="submit">Lọc</button></form>` : ""}
    <div id="mapScreenText" role="tabpanel">${content}</div></div>`;
}

function renderMap() {
  if (G.resizeObserver) { G.resizeObserver.disconnect(); G.resizeObserver = null; }
  const ws = mapWorkspace(), data = M.data, meta = data && data.meta;
  const workspaceSelect = `<select id="mapWorkspace" aria-label="Workspace sitemap"><option value="">Chọn workspace</option>${S.workspaces.map(workspace => `<option value="${esc(workspace.name)}" ${workspace.name === ws ? "selected" : ""}>${esc(workspace.name)}</option>`).join("")}</select>`;
  let body;
  if (!ws) body = '<div class="diagram-empty"><h2>Chọn workspace để khám phá</h2><p>Sơ đồ hiển thị cách menu, action và view nối với nhau.</p></div>';
  else if (M.loading) body = '<div class="diagram-empty" role="status"><h2>Đang dựng sơ đồ…</h2><p>Đang đọc sitemap của workspace.</p></div>';
  else if (M.error) body = `<div class="map-error" role="alert">${esc(M.error)}</div>`;
  else if (data && !data.available) body = `<div class="diagram-empty"><h2>Workspace chưa có sitemap</h2><p>Nhờ AI dựng sitemap cho <b>${esc(ws)}</b>, rồi bấm tải lại.</p></div>`;
  else if (data) {
    G.layout = graphLayout();
    const focus = G.layout.index.nodes.get(G.focus), bounds = G.layout.bounds;
    const miniNodes = [...G.layout.positions.values()].map(node => `<rect x="${node.x}" y="${node.y}" width="280" height="96" rx="10" fill="${node.kind === "action" ? "var(--map-purple)" : node.kind === "view" ? "var(--map-teal)" : "var(--map-blue)"}" opacity=".7"/>`).join("");
    const search = G.searchOpen && M.results !== null ? `<section class="map-results"><header>${M.results.length} kết quả<button class="map-close" id="mapCloseSearch" aria-label="Đóng kết quả">×</button></header>${M.results.map(item => `<button class="map-result" data-map-screen="${esc(item.key)}"><strong>${esc(item.label)}</strong><small>${esc(item.kind)} · ${esc(item.key)}</small></button>`).join("") || '<p class="map-note">Không có kết quả. Thử nhãn Anh/Nhật hoặc tên kỹ thuật.</p>'}${M.results.length === 60 ? '<p class="map-note">Hiển thị tối đa 60 kết quả.</p>' : ""}</section>` : "";
    body = `<div class="diagram-space"><div class="diagram-breadcrumb"><button data-map-home="true">Tổng quan workspace</button>${focus ? `<span> / </span><strong>${esc(focus.label)}</strong>` : '<strong>Ứng dụng & menu gốc</strong>'}</div>
      <svg id="mapDiagram" tabindex="0" role="img" aria-label="Sơ đồ menu, action và view của workspace">${graphSvg(G.layout)}</svg>${search}
      <div class="diagram-legend"><span><i></i>Menu</span><span class="action"><i></i>Action</span><span class="view"><i></i>View</span></div>
      <div class="diagram-tools"><button id="mapZoomOut" aria-label="Thu nhỏ">−</button><output id="mapZoomLevel" aria-label="Mức thu phóng">100%</output><button id="mapZoomIn" aria-label="Phóng to">+</button><button id="mapFit">Vừa khung</button><button id="mapCollapse" ${G.focus ? "" : "disabled"}>Thu nhánh</button></div>
      <svg class="diagram-minimap" viewBox="${bounds.x} ${bounds.y} ${bounds.width} ${bounds.height}" aria-hidden="true">${miniNodes}<rect id="mapMiniViewport" fill="none" stroke="var(--map-blue)" stroke-width="8"/></svg></div>
      ${G.inspector ? `<aside class="map-inspector" id="mapScreen" aria-label="Chi tiết màn hình">${renderMapScreen()}</aside>` : ""}`;
  }
  $("siteMapMain").innerHTML = `<header class="map-heading"><div><div class="map-eyebrow">WORKSPACE EXPLORER</div><h1>Sơ đồ workspace</h1></div>${meta ? `<div class="map-stats"><span><strong>${meta.menus}</strong> menu</span><span><strong>${meta.actions}</strong> action</span><span><strong>${meta.views}</strong> view</span></div>` : ""}</header>
    <form id="mapSearchForm" class="map-toolbar"><label>${workspaceSelect}</label><label class="map-query"><span class="map-search-box"><input type="search" id="mapQuery" value="${esc(M.q)}" placeholder="Tìm menu, action, view…  English / 日本語" aria-label="Tìm trong sitemap"><button class="btn primary" type="submit">Tìm</button></span></label>
    <label><select class="map-kind" id="mapKind" aria-label="Loại màn hình"><option value="">Tất cả</option>${["menu", "action", "view"].map(kind => `<option ${M.kind === kind ? "selected" : ""}>${kind}</option>`).join("")}</select></label><label><select id="mapLanguage" aria-label="Ngôn ngữ nội dung"><option value="en" ${M.lang !== "ja" ? "selected" : ""}>English</option><option value="ja" ${M.lang === "ja" ? "selected" : ""}>日本語</option></select></label><button class="btn" type="button" id="mapReload" ${M.loading ? "disabled" : ""} aria-label="Tải lại sitemap">↻</button></form>
    <section class="diagram-shell${G.inspector ? " has-inspector" : ""}">${body}</section><footer class="diagram-footer"><span>Kéo để di chuyển · cuộn để zoom · dấu + để mở nhánh · bấm node để xem chi tiết</span><span>${meta ? `${G.layout ? G.layout.positions.size : 0} node đang hiển thị · Odoo ${esc(meta.odoo_version)} · ${esc(meta.generated.slice(0,10))}${data.unresolved_count ? " · " + data.unresolved_count + " mục chưa đọc được" : ""}` : ""}</span></footer>`;
  const svg = $("mapDiagram");
  if (svg) {
    if (G.fit || (G.inspector && window.innerWidth <= 1100)) { graphFit(true); G.fit = false; } else graphTransform();
    let measured = `${svg.clientWidth}|${svg.clientHeight}`;
    G.resizeObserver = new ResizeObserver(() => {
      const size = `${svg.clientWidth}|${svg.clientHeight}`;
      if (size !== measured && svg.isConnected && S.tab === "site_map") { measured = size; graphFit(true); }
    });
    G.resizeObserver.observe(svg);
    svg.addEventListener("wheel", event => { event.preventDefault(); const bounds = svg.getBoundingClientRect(); graphZoom(event.deltaY < 0 ? 1.12 : 1/1.12, event.clientX - bounds.left, event.clientY - bounds.top); }, { passive: false });
    svg.addEventListener("pointerdown", event => {
      if (event.target.closest("[role=button]") || event.button !== 0) return;
      G.drag = { id: event.pointerId, x: event.clientX, y: event.clientY, ox: G.x, oy: G.y }; svg.setPointerCapture(event.pointerId); svg.classList.add("dragging");
    });
    svg.addEventListener("pointermove", event => { if (G.drag && G.drag.id === event.pointerId) { G.x = G.drag.ox + event.clientX - G.drag.x; G.y = G.drag.oy + event.clientY - G.drag.y; graphTransform(); } });
    const release = () => { G.drag = null; svg.classList.remove("dragging"); };
    svg.addEventListener("pointerup", release); svg.addEventListener("pointercancel", release);
  }
}

async function loadMap() {
  const ws = mapWorkspace();
  if (ws !== M.ws) {
    Object.assign(M, { ws, lang: (S.workspaces.find(w => w.name === ws) || {}).language === "ja" ? "ja" : "en", q: "", kind: "", key: "", view: "", match: "", required: false, data: null, results: null, screen: null, detailError: "" });
    G.focus = ""; G.expanded.clear(); G.inspector = false; G.route = []; G.searchOpen = false; G.fit = true;
    M.detailRequest++;
  }
  const request = ++M.request;
  M.detailRequest++; M.detailLoading = false;
  if (!ws) return renderMap();
  M.loading = true; M.error = ""; renderMap();
  try {
    const data = await mapGet({ what: "show" });
    if (request !== M.request) return;
    let results = null;
    if (data.available && (M.q.trim() || M.kind)) { const found = await mapGet({ what: "find", q: M.q, kind: M.kind }); results = found.results || []; }
    if (request !== M.request) return;
    M.data = data; M.results = results; M.loading = false; renderMap();
    if (data.available && M.key && G.inspector) loadMapScreen();
  } catch (error) { if (request === M.request) { M.loading = false; M.error = error.message; renderMap(); } }
}

async function loadMapScreen() {
  const request = ++M.detailRequest;
  M.detailLoading = true; M.detailError = ""; renderMap();
  try {
    const screen = await mapGet({ what: "screen", screen: M.key, view: M.view, match: M.match, required: String(M.required) });
    if (request !== M.detailRequest) return;
    M.screen = screen; M.detailLoading = false; G.route = screen.route_edges || []; renderMap();
  } catch (error) { if (request === M.detailRequest) { M.detailLoading = false; M.detailError = error.message; renderMap(); } }
}

window.addEventListener("resize", () => { if (S.tab === "site_map" && $("mapDiagram")) graphFit(true); });
