/**
 * The test dashboard: one self-contained page served at /tests. It only reads: runs, live progress, each case's steps and pictures, and what the
 * assistant has learned. Every piece of text that comes from a run goes in through textContent, never as markup.
 *
 * Design: a monitor for someone who wants three answers at a glance: what is happening now, what needs me, and where is the evidence.
 * Neutral surfaces, one accent, status colour plus a shape (never colour alone), numbers instead of unlabeled bars, dark mode, no motion beyond
 * one small live indicator.
 */
export function dashboardPage(token: string): string {
  return `<!doctype html>
<html lang="vi"><head><meta charset="utf-8" /><meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Kiểm thử · Office Partner</title>
<style>
:root{
  --bg:#f4f5f7;--surface:#fff;--sunken:#f0f2f5;--ink:#15181d;--ink-2:#4a5260;--ink-3:#6b7382;--line:#e2e5ea;--line-2:#cfd4db;
  --accent:#2250c7;--accent-soft:#e8eefc;
  --pass:#0f7a56;--pass-soft:#e3f4ec;--fail:#c1272d;--fail-soft:#fbe9ea;--warn:#a35f00;--warn-soft:#fdf1dc;--run:#2250c7;--run-soft:#e8eefc;--none:#6b7382;--plan:#a3abb8;
  --radius:10px;--radius-s:6px;
  --font:ui-sans-serif,system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;--mono:ui-monospace,"Cascadia Mono","SF Mono",Menlo,Consolas,monospace;
}
@media (prefers-color-scheme:dark){:root{
  --bg:#0e1116;--surface:#161a21;--sunken:#1c212a;--ink:#e9ecf1;--ink-2:#b4bbc7;--ink-3:#8b93a2;--line:#262c36;--line-2:#363d4a;
  --accent:#7aa2ff;--accent-soft:#1c2740;
  --pass:#4cc38a;--pass-soft:#14291f;--fail:#ff7d82;--fail-soft:#31181a;--warn:#e6a94f;--warn-soft:#2e2412;--run:#7aa2ff;--run-soft:#1c2740;--none:#8b93a2;--plan:#4a5262;
}}
*{box-sizing:border-box}
[hidden]{display:none!important}
html,body{height:100%}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 var(--font);-webkit-font-smoothing:antialiased}
button{font:inherit;color:inherit;cursor:pointer}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px;border-radius:4px}
.num{font-variant-numeric:tabular-nums}
.mono{font-family:var(--mono)}

/* header */
header.top{position:sticky;top:0;z-index:5;display:flex;align-items:center;gap:20px;height:52px;padding:0 24px;background:var(--surface);border-bottom:1px solid var(--line)}
.brand{font-weight:650;font-size:15px;letter-spacing:-.01em}
.tabs{display:flex;gap:4px;height:100%}
.tabs button{position:relative;border:0;background:none;padding:0 12px;color:var(--ink-3);font-weight:550}
.tabs button[aria-selected=true]{color:var(--ink)}
.tabs button[aria-selected=true]::after{content:"";position:absolute;left:12px;right:12px;bottom:-1px;height:2px;background:var(--accent);border-radius:2px}
.live{margin-left:auto;display:flex;align-items:center;gap:8px;color:var(--ink-3);font-size:12.5px}
.dot{width:8px;height:8px;border-radius:50%;background:var(--plan);flex:none}
.dot.on{background:var(--pass)}
.dot.pulse{background:var(--run);animation:pulse 1.6s ease-in-out infinite}
.dot.off{background:var(--fail)}
@keyframes pulse{0%,100%{box-shadow:0 0 0 0 color-mix(in srgb,var(--run) 45%,transparent)}50%{box-shadow:0 0 0 5px transparent}}
@media (prefers-reduced-motion:reduce){.dot.pulse{animation:none}.ring{animation:none!important}}

/* layout */
.shell{display:grid;grid-template-columns:300px minmax(0,1fr);gap:0;height:calc(100% - 52px)}
.rail{overflow:auto;border-right:1px solid var(--line);background:var(--surface);padding:12px 8px 24px}
.main{overflow:auto;padding:24px 28px 56px}
.rail h2{margin:12px 12px 6px;font-size:11.5px;font-weight:650;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3)}
.runbtn{display:block;width:100%;text-align:left;border:0;background:none;border-radius:var(--radius-s);padding:10px 12px;margin:1px 0}
.runbtn:hover{background:var(--sunken)}
.runbtn[aria-current=true]{background:var(--accent-soft)}
.runbtn .l1{display:flex;align-items:center;gap:8px}
.runbtn .name{flex:1;min-width:0;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.runbtn .time{color:var(--ink-3);font-size:12px}
.runbtn .l2{margin:2px 0 0 26px;color:var(--ink-3);font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.seg{display:flex;gap:2px;height:4px;margin:7px 0 0 26px;border-radius:3px;overflow:hidden;background:var(--line)}
.seg i{display:block;height:100%}
.empty{color:var(--ink-3);padding:24px 12px}
.empty b{display:block;color:var(--ink);margin-bottom:4px}

/* status */
.ico{width:16px;height:16px;flex:none;display:inline-block;vertical-align:-3px}
.s-pass{color:var(--pass)}.s-fail{color:var(--fail)}.s-run{color:var(--run)}.s-warn{color:var(--warn)}.s-none{color:var(--none)}.s-plan{color:var(--plan)}
.ring{animation:spin 1.1s linear infinite;transform-origin:50% 50%}
@keyframes spin{to{transform:rotate(360deg)}}
.badge{display:inline-flex;align-items:center;gap:6px;padding:2px 10px 2px 8px;border-radius:999px;font-size:12.5px;font-weight:600;white-space:nowrap}
.badge.s-pass{background:var(--pass-soft)}.badge.s-fail{background:var(--fail-soft)}.badge.s-run{background:var(--run-soft)}.badge.s-warn{background:var(--warn-soft)}.badge.s-none,.badge.s-plan{background:var(--sunken);color:var(--ink-2)}

/* run header */
.runhead{display:flex;align-items:flex-start;gap:12px;flex-wrap:wrap}
.runhead h1{margin:0;font-size:20px;line-height:1.25;letter-spacing:-.015em}
.runmeta{margin:6px 0 0;color:var(--ink-3);font-size:13px;display:flex;gap:6px 14px;flex-wrap:wrap}
.runmeta span+span::before{content:"·";margin-right:14px;color:var(--line-2)}

.stats{display:flex;flex-wrap:wrap;gap:8px 28px;margin:20px 0 10px}
.stat{display:flex;align-items:baseline;gap:8px}
.stat b{font-size:26px;line-height:1;letter-spacing:-.02em}
.stat span{color:var(--ink-2);font-size:13px}
.bar{display:flex;gap:2px;height:10px;border-radius:5px;overflow:hidden;background:var(--line)}
.bar i{display:block;height:100%;min-width:3px}

/* attention + now */
.note{margin:20px 0 0;padding:14px 16px;border-radius:var(--radius);border:1px solid var(--line);background:var(--surface);display:flex;gap:12px}
.note.attn{border-color:color-mix(in srgb,var(--fail) 40%,var(--line));background:var(--fail-soft)}
.note.ask{border-color:color-mix(in srgb,var(--warn) 45%,var(--line));background:var(--warn-soft)}
.note .body{min-width:0;flex:1}
.note h3{margin:0 0 2px;font-size:14px}
.note p{margin:0;color:var(--ink-2)}
.note details{margin-top:8px;color:var(--ink-3);font-size:12.5px}
.note summary{cursor:pointer}
.note pre{margin:6px 0 0;white-space:pre-wrap;overflow-wrap:anywhere;font:12px/1.5 var(--mono);color:var(--ink-2)}
.now{display:flex;gap:14px;align-items:center}
.now .thumb{width:148px;flex:none;border-radius:var(--radius-s);border:1px solid var(--line);overflow:hidden;background:var(--sunken)}
.now .thumb img{display:block;width:100%;height:auto}
.now .k{font-size:11.5px;font-weight:650;letter-spacing:.06em;text-transform:uppercase;color:var(--run)}

/* cases + detail */
.split{display:grid;grid-template-columns:minmax(300px,5fr) minmax(0,7fr);gap:20px;margin-top:24px;align-items:start}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius)}
.panel>h2{margin:0;padding:12px 16px;font-size:12px;font-weight:650;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3);border-bottom:1px solid var(--line)}
.row{display:grid;grid-template-columns:20px 44px minmax(0,1fr) auto;gap:10px;align-items:start;width:100%;text-align:left;border:0;border-bottom:1px solid var(--line);background:none;padding:11px 16px}
.row:last-child{border-bottom:0}
.row:hover{background:var(--sunken)}
.row[aria-current=true]{background:var(--accent-soft);box-shadow:inset 3px 0 0 var(--accent)}
.row .id{color:var(--ink-3);font-size:12.5px;padding-top:1px}
.row .t{min-width:0}
.row .t b{display:block;font-weight:550;overflow-wrap:anywhere}
.row .t small{display:block;color:var(--ink-3);font-size:12.5px;margin-top:1px;overflow-wrap:anywhere}
.row .r{color:var(--ink-3);font-size:12.5px;white-space:nowrap;padding-top:1px}
.detail{position:sticky;top:0;max-height:calc(100vh - 100px);overflow:auto}
.dhead{padding:16px 18px;border-bottom:1px solid var(--line)}
.dhead h2{margin:0 0 6px;font-size:16px;line-height:1.3;letter-spacing:-.01em;text-transform:none;font-weight:650;color:var(--ink);padding:0;border:0}
.kv{margin:8px 0 0;color:var(--ink-2)}
.kv b{display:block;font-size:11.5px;font-weight:650;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3);margin-bottom:1px}
.steps{list-style:none;margin:0;padding:6px 0}
.step{display:grid;grid-template-columns:20px minmax(0,1fr);gap:10px;padding:10px 18px}
.step+.step{border-top:1px dashed var(--line)}
.step .plan{color:var(--ink-3);font-size:12.5px;overflow-wrap:anywhere}
.step .did{font-weight:550;overflow-wrap:anywhere}
.step .act{color:var(--ink-2);font-size:13px;overflow-wrap:anywhere}
.step .why{margin-top:6px;padding:8px 10px;border-radius:var(--radius-s);background:var(--fail-soft);color:var(--ink);font-size:12.5px;overflow-wrap:anywhere}
.step .todo{color:var(--ink-3);font-style:italic}
.evh{display:flex;align-items:center;justify-content:space-between;padding:12px 18px 8px;border-top:1px solid var(--line)}
.evh h3{margin:0;font-size:12px;font-weight:650;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3)}
.sb{min-width:0}
.thumbs{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}
.thumbs button{position:relative;display:block;width:132px;padding:0;border:1px solid var(--line);border-radius:var(--radius-s);overflow:hidden;background:var(--sunken)}
.thumbs button:hover{border-color:var(--accent)}
.thumbs img{display:block;width:100%;height:auto}
.thumbs .k{position:absolute;left:5px;top:5px;padding:0 6px;border-radius:999px;background:rgba(0,0,0,.62);color:#fff;font-size:11px}
.thumbs .new{position:absolute;right:5px;top:5px;padding:0 6px;border-radius:999px;background:var(--run);color:#fff;font-size:11px;font-weight:600}
.none-sel{padding:40px 24px;color:var(--ink-3);text-align:center}

/* timeline */
.tl{margin-top:24px}
.tl summary{cursor:pointer;color:var(--ink-3);font-size:13px;padding:6px 0}
.tl div{display:grid;grid-template-columns:70px 110px 1fr;gap:10px;padding:5px 0;border-top:1px solid var(--line);font-size:13px}
.tl .ts{color:var(--ink-3)}.tl .ph{font-weight:600}

/* knowledge */
.know{max-width:980px;margin:0 auto}
.know .tools{display:flex;gap:10px;margin:0 0 16px}
.know input,.know select{font:inherit;color:inherit;background:var(--surface);border:1px solid var(--line-2);border-radius:var(--radius-s);padding:8px 12px}
.know input{flex:1}
.k-group{margin:22px 0 8px;font-size:12px;font-weight:650;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-3)}
.k-item{display:grid;grid-template-columns:96px minmax(0,1fr);gap:14px;padding:10px 14px;background:var(--surface);border:1px solid var(--line);border-radius:var(--radius-s);margin-bottom:6px}
.k-item .kind{font-size:12px;font-weight:650;color:var(--accent)}
.k-item .tags{color:var(--ink-3);font-size:12.5px;margin-top:2px}

/* lightbox */
.lb{position:fixed;inset:0;z-index:20;background:rgba(8,10,14,.95);display:none;flex-direction:column}
.lb.open{display:flex}
.lb .bar2{display:flex;align-items:center;gap:14px;padding:12px 18px;color:#e9ecf1;font-size:13px}
.lb .bar2 .sp{flex:1}
.lb .bar2 a,.lb .bar2 button{color:#e9ecf1;background:rgba(255,255,255,.1);border:0;border-radius:var(--radius-s);padding:6px 12px;text-decoration:none}
.lb .stage{flex:1;min-height:0;display:flex;align-items:center;justify-content:center;padding:0 18px 12px}
.lb .strip{display:flex;gap:8px;overflow-x:auto;padding:0 18px 16px}
.lb .strip button{flex:none;width:96px;padding:0;border:2px solid transparent;border-radius:6px;overflow:hidden;background:#222;opacity:.7}
.lb .strip button[aria-current=true]{border-color:#7aa2ff;opacity:1}
.lb .strip img{display:block;width:100%;height:auto}
.lb img{max-width:100%;max-height:100%;border-radius:var(--radius-s);background:#fff}

@media (max-width:1100px){.split{grid-template-columns:1fr}.detail{position:static;max-height:none}}
@media (max-width:860px){.shell{grid-template-columns:1fr;height:auto}.rail{max-height:240px;border-right:0;border-bottom:1px solid var(--line)}.main{padding:18px 16px 48px}header.top{padding:0 14px;gap:10px}.live span.t{display:none}}
</style></head><body>
<header class="top">
  <div class="brand">Kiểm thử</div>
  <nav class="tabs" role="tablist"><button id="tab-run" role="tab" aria-selected="true">Lần chạy</button><button id="tab-know" role="tab" aria-selected="false">Tri thức đã học</button></nav>
  <div class="live" role="status" aria-live="polite"><i id="live-dot" class="dot"></i><span id="live-text" class="t">Đang kết nối</span></div>
</header>
<div id="view-run" class="shell">
  <aside class="rail" id="rail" aria-label="Các lần chạy"></aside>
  <main class="main" id="main"></main>
</div>
<div id="view-know" class="main" hidden><div class="know" id="know"></div></div>
<div class="lb" id="lb" role="dialog" aria-modal="true" aria-label="Ảnh bằng chứng"><div class="bar2"><span id="lb-cap"></span><span class="sp"></span><span id="lb-count" class="num"></span><a id="lb-open" target="_blank" rel="noopener">Mở ảnh gốc</a><button id="lb-prev" type="button" aria-label="Ảnh trước (←)">←</button><button id="lb-next" type="button" aria-label="Ảnh sau (→)">→</button><button id="lb-close" type="button">Đóng (Esc)</button></div><div class="stage"><img id="lb-img" alt="Ảnh bằng chứng" /></div><div class="strip" id="lb-strip" aria-label="Các ảnh của case"></div></div>
<script>
(function(){
var T=${JSON.stringify(token)};
var NS="http://www.w3.org/2000/svg";
var $=function(id){return document.getElementById(id)};
function h(tag,cls,kids){var e=document.createElement(tag);if(cls)e.className=cls;(kids||[]).forEach(function(k){if(k==null||k===false)return;e.appendChild(typeof k==="string"?document.createTextNode(k):k)});return e}
function api(p,q){return fetch("/tests/api/"+p+"?t="+encodeURIComponent(T)+(q||"")).then(function(r){if(!r.ok)throw new Error(r.status);return r.json()})}
function imgUrl(id,p){return "/tests/api/image?t="+encodeURIComponent(T)+"&id="+encodeURIComponent(id)+"&path="+encodeURIComponent(p)}

/* icons: a shape per status, so colour is never the only signal */
var PATHS={check:"M3.5 8.5l3 3 6-6.5",x:"M4.5 4.5l7 7M11.5 4.5l-7 7",pause:"M6 4.5v7M10 4.5v7",dash:"M4.5 8h7",dot:"M8 8h.01"};
function icon(name,cls){
  var s=document.createElementNS(NS,"svg");s.setAttribute("viewBox","0 0 16 16");s.setAttribute("class","ico "+(cls||""));s.setAttribute("aria-hidden","true");
  if(name==="ring"){var c=document.createElementNS(NS,"circle");c.setAttribute("cx","8");c.setAttribute("cy","8");c.setAttribute("r","5.5");c.setAttribute("fill","none");c.setAttribute("stroke","currentColor");c.setAttribute("stroke-width","2");c.setAttribute("stroke-dasharray","20 15");c.setAttribute("stroke-linecap","round");c.setAttribute("class","ring");s.appendChild(c);return s}
  var p=document.createElementNS(NS,"path");p.setAttribute("d",PATHS[name]);p.setAttribute("fill","none");p.setAttribute("stroke","currentColor");p.setAttribute("stroke-width",name==="dot"?"4":"2");p.setAttribute("stroke-linecap","round");p.setAttribute("stroke-linejoin","round");s.appendChild(p);return s}

var ST={PASS:{vi:"Đạt",ic:"check",c:"s-pass",col:"var(--pass)"},FAIL:{vi:"Lỗi",ic:"x",c:"s-fail",col:"var(--fail)"},RUNNING:{vi:"Đang chạy",ic:"ring",c:"s-run",col:"var(--run)"},
 PENDING:{vi:"Chờ người",ic:"pause",c:"s-warn",col:"var(--warn)"},BLOCKED:{vi:"Bị chặn",ic:"pause",c:"s-warn",col:"var(--warn)"},PLANNED:{vi:"Chưa chạy",ic:"dot",c:"s-plan",col:"var(--plan)"},
 "NOT EXECUTED":{vi:"Chưa thực hiện",ic:"dash",c:"s-none",col:"var(--none)"},"NOT AUTOMATED":{vi:"Chưa tự động hóa",ic:"dash",c:"s-none",col:"var(--none)"}};
var ORDER=["PASS","FAIL","RUNNING","PENDING","BLOCKED","PLANNED","NOT EXECUTED","NOT AUTOMATED"];
function outcome(s){
  if(s.phase!=="finished")return PHASE[s.phase]||PHASE.idle;
  var c=s.counts;if(c.FAIL)return{vi:"Đã xong · có case lỗi",ic:"x",c:"s-fail"};
  if(c.PENDING||c.BLOCKED||c["NOT EXECUTED"])return{vi:"Đã xong · còn việc cần xem",ic:"pause",c:"s-warn"};
  return{vi:"Đã xong · tất cả đạt",ic:"check",c:"s-pass"}}
var PHASE={preparing:{vi:"Đang chuẩn bị",ic:"ring",c:"s-run"},planned:{vi:"Đã lập kế hoạch",ic:"dot",c:"s-plan"},running:{vi:"Đang chạy",ic:"ring",c:"s-run"},paused:{vi:"Tạm dừng để sửa",ic:"pause",c:"s-fail"},
 finished:{vi:"Đã xong",ic:"check",c:"s-pass"},stopped:{vi:"Tự dừng",ic:"x",c:"s-fail"},idle:{vi:"Không hoạt động",ic:"dash",c:"s-none"}};
var STAGE={setup:"Thiết lập",learn:"Tìm hiểu hệ thống",automate:"Viết automation",trial:"Chạy thử",preflight:"Kiểm tra môi trường",plan:"Lập kế hoạch",run:"Đang chạy test",fix:"Đang sửa lỗi",cleanup:"Dọn dữ liệu test",report:"Lập báo cáo",done:"Hoàn tất",blocked:"Cần bạn"};
function st(s){return ST[s]||ST["NOT EXECUTED"]}
function badge(s){var m=st(s);var b=h("span","badge "+m.c,[icon(m.ic),m.vi]);b.title=s;return b}

/* a run id is YYYYMMDD_HHMMSS_<project>_<label>, adhoc-YYYYMMDD, or something typed by hand */
function parseId(id){
  var m=/^(\\d{4})(\\d{2})(\\d{2})_(\\d{2})(\\d{2})(\\d{2})_[^_]*_?(.*)$/.exec(id);
  if(m){var label=(m[7]||"").replace(/-/g," ").trim();return{day:m[1]+"-"+m[2]+"-"+m[3],time:m[4]+":"+m[5],name:label?label.charAt(0).toUpperCase()+label.slice(1):"Lần chạy"}}
  var a=/^adhoc-(\\d{4})(\\d{2})(\\d{2})$/.exec(id);if(a)return{day:a[1]+"-"+a[2]+"-"+a[3],time:"",name:"Chạy thử trong ngày"};
  return{day:"",time:"",name:id.replace(/^adhoc-/,"Chạy thử ").replace(/_/g," ")}}
function today(){var d=new Date();return d.getFullYear()+"-"+("0"+(d.getMonth()+1)).slice(-2)+"-"+("0"+d.getDate()).slice(-2)}
function dayLabel(d){return d===today()?"Hôm nay":d||"Khác"}
function ts(s){return s?Date.parse(s.replace(" ","T")):NaN}
function duration(a,b){var ms=ts(b)-ts(a);if(!(ms>=0))return"";var s=Math.round(ms/1000);if(s<60)return s+" giây";var m=Math.floor(s/60);return m+" phút "+("0"+(s%60)).slice(-2)+" giây"}
function summary(c){var parts=[];ORDER.forEach(function(k){if(c[k]&&k!=="PLANNED"&&k!=="NOT AUTOMATED")parts.push(c[k]+" "+st(k).vi.toLowerCase())});if(c.PLANNED)parts.push(c.PLANNED+" chưa chạy");return parts.join(" · ")||"chưa có kết quả"}
function seg(c,cls){var t=0,k;for(k in c)t+=c[k];var bar=h("div",cls||"seg");if(!t)return bar;ORDER.forEach(function(s){if(c[s]){var i=h("i");i.style.width=(100*c[s]/t)+"%";i.style.background=st(s).col;i.title=st(s).vi+": "+c[s];bar.appendChild(i)}});return bar}

var cur=null,runs=[],state=null,sel=null,tab="run",lastRuns="",lastState="",lastOk=Date.now();

/* live indicator */
function live(){var d=$("live-dot"),t=$("live-text"),run=runs.filter(function(r){return r.phase==="running"||r.phase==="preparing"}).length;
  if(Date.now()-lastOk>6000){d.className="dot off";t.textContent="Mất kết nối — tải lại trang";return}
  d.className=run?"dot pulse":"dot on";t.textContent=(run?"Đang theo dõi trực tiếp":"Đã kết nối")+" · "+new Date().toLocaleTimeString("vi-VN")}

/* rail */
/* a redraw rebuilds the DOM; this puts back what the person was doing: scroll position (page, panel, detail), open sections, focused item */
function keepView(box,draw){
  var det=box.querySelector(".detail"),snap={s:box.scrollTop,d:det?det.scrollTop:0,w:window.scrollY,open:{}};
  box.querySelectorAll("details[data-k]").forEach(function(d){snap.open[d.getAttribute("data-k")]=d.open});
  var fk=document.activeElement&&document.activeElement.getAttribute&&document.activeElement.getAttribute("data-fk");
  var held=box.offsetHeight;box.style.minHeight=held+"px";
  draw();
  box.querySelectorAll("details[data-k]").forEach(function(d){var k=d.getAttribute("data-k");if(snap.open[k]!==undefined)d.open=snap.open[k]});
  box.scrollTop=snap.s;var det2=box.querySelector(".detail");if(det2)det2.scrollTop=snap.d;window.scrollTo(0,snap.w);
  box.style.minHeight="";
  if(fk){var el=box.querySelector('[data-fk="'+fk.replace(/"/g,'\\"')+'"]');if(el&&el.focus)el.focus({preventScroll:true})}}
function renderRail(){keepView($("rail"),drawRail)}
function drawRail(){var box=$("rail");box.textContent="";
  if(!runs.length){box.appendChild(h("div","empty",[h("b",0,["Chưa có lần chạy nào"]),"Nhắn /system-test trong add-in Office Partner để bắt đầu. Kế hoạch và tiến độ sẽ hiện ở đây."]));return}
  var sorted=runs.slice().sort(function(a,b){return (ts(b.started)||0)-(ts(a.started)||0)});var group=null;
  sorted.forEach(function(r){var p=parseId(r.id);var g=dayLabel(p.day);if(g!==group){group=g;box.appendChild(h("h2",0,[g]))}
    var ph=outcome(r);var b=h("button","runbtn",[h("div","l1",[icon(ph.ic,ph.c),h("span","name",[p.name]),h("span","time num",[p.time||""])]),h("div","l2",[ph.vi+" · "+summary(r.counts)]),seg(r.counts)]);
    b.setAttribute("aria-current",String(r.id===cur));b.setAttribute("data-fk","run:"+r.id);b.title=r.id;b.onclick=function(){cur=r.id;sel=null;state=null;lastState="";renderRail();refreshRun()};box.appendChild(b)})}

function tech(text){var d=h("details",0,[h("summary",0,["Chi tiết kỹ thuật"]),h("pre",0,[text])]);d.setAttribute("data-k","tech");return d}
/* the thing that needs a person, in plain words first; the raw text stays one click away */
function attention(s){
  if(s.phase==="paused"){var m=/case (\\S+) needs repair/.exec(s.stopReason||"");var left=s.cases.filter(function(c){return c.status==="PLANNED"}).length;
    return h("div","note attn",[icon("pause","s-fail"),h("div","body",[h("h3",0,["Tạm dừng để sửa"+(m?" case "+m[1]:"")]),h("p",0,["Case này cần sửa trước khi chạy tiếp."+(left?" "+left+" case chưa tới được giữ nguyên.":"")]),tech(s.stopReason)])])}
  if(s.phase==="stopped")return h("div","note attn",[icon("x","s-fail"),h("div","body",[h("h3",0,["Run tự dừng vì cùng một lỗi lặp lại"]),h("p",0,["Các case chưa chạy được ghi là Chưa thực hiện."]),tech(s.stopReason)])]);
  var ask=s.phases.filter(function(p){return p.phase==="blocked"&&/^ASK THE USER/.test(p.text)}).pop();
  if(ask&&s.phase!=="finished")return h("div","note ask",[icon("pause","s-warn"),h("div","body",[h("h3",0,["Cần bạn trả lời"]),h("p",0,[ask.text.replace(/^ASK THE USER:\\s*/,"")])])]);
  return null}
function nowStrip(s){
  if(s.phase!=="running"&&s.phase!=="preparing")return null;
  var cp=s.currentPhase,running=s.cases.filter(function(c){return c.status==="RUNNING"})[0];var kids=[];
  var img=running&&running.images.length?running.images[running.images.length-1]:null;
  if(img){var th=h("div","thumb");var i=h("img");i.alt="Ảnh mới nhất";i.src=imgUrl(s.id,img);th.appendChild(i);kids.push(th)}
  var last=running&&running.steps.length?running.steps[running.steps.length-1]:null;
  kids.push(h("div","body",[h("div","k",["Đang diễn ra"]),h("h3",0,[running?"Case "+running.id+(running.title?" · "+running.title:""):(cp?STAGE[cp.phase]||cp.phase:"Đang chuẩn bị")]),
    h("p",0,[running?(last?"Bước "+running.step+(running.total?"/"+running.total:"")+": "+last.caption:"Bắt đầu case"):(cp?cp.text:"")])]));
  return h("div","note now",kids)}

/* main */
function renderRun(){keepView($("main"),drawRun)}
function drawRun(){var box=$("main");box.textContent="";
  if(!state){box.appendChild(h("div","empty",[runs.length?"Chọn một lần chạy ở bên trái.":""]));return}
  var s=state,p=parseId(s.id),ph=outcome(s);
  box.appendChild(h("div","runhead",[h("h1",0,[p.name]),h("span","badge "+ph.c,[icon(ph.ic),ph.vi])]));
  var meta=[];if(p.day)meta.push((p.day===today()?"Hôm nay":p.day)+(p.time?" "+p.time:""));var started=s.phase!=="planned"&&s.phase!=="preparing"&&s.phase!=="idle";var live2=s.phase==="running";var du=started?duration(s.started,live2?new Date(Date.now()-new Date().getTimezoneOffset()*60000).toISOString().slice(0,19).replace("T"," "):s.updated):"";if(du)meta.push(du);meta.push(s.cases.length+" case");
  var mm=h("div","runmeta");meta.forEach(function(x){mm.appendChild(h("span",0,[x]))});var idm=h("span","mono",[s.id]);idm.style.opacity=".7";mm.appendChild(idm);box.appendChild(mm);
  var stats=h("div","stats");ORDER.forEach(function(k){if(s.counts[k]){var m=st(k);var one=h("div","stat",[h("b","num "+m.c,[String(s.counts[k])]),h("span",0,[m.vi])]);stats.appendChild(one)}});box.appendChild(stats);box.appendChild(seg(s.counts,"bar"));
  var a=attention(s);if(a)box.appendChild(a);var n=nowStrip(s);if(n)box.appendChild(n);
  var split=h("div","split");split.appendChild(casesPanel(s));split.appendChild(detailPanel(s));box.appendChild(split);
  if(s.phases.length){var tl=h("details","tl",[h("summary",0,["Nhật ký các giai đoạn ("+s.phases.length+")"])]);tl.setAttribute("data-k","tl");s.phases.slice().reverse().forEach(function(q){tl.appendChild(h("div",0,[h("span","ts num",[q.ts.slice(11)]),h("span","ph",[STAGE[q.phase]||q.phase]),h("span",0,[q.text])]))});box.appendChild(tl)}}

function key(c){return c.suite+"/"+c.id}
function pickDefault(s){var r=s.cases.filter(function(c){return c.status==="RUNNING"})[0]||s.cases.filter(function(c){return c.status==="FAIL"||c.status==="BLOCKED"||c.status==="NOT EXECUTED"})[0]||s.cases[0];return r?key(r):null}
function casesPanel(s){var panel=h("section","panel",[h("h2",0,["Các case ("+s.cases.length+")"])]);
  s.cases.forEach(function(c){var m=st(c.status);var right=c.total?(c.status==="PLANNED"?c.total+" bước":c.step+"/"+c.total):"";
    var t=h("div","t",[h("b",0,[c.title||("Case "+c.id)]),c.remark&&c.status!=="PASS"?h("small",0,[c.remark.length>110?c.remark.slice(0,110)+"…":c.remark]):null]);
    var b=h("button","row",[icon(m.ic,m.c),h("span","id mono",[c.id]),t,h("span","r num",[right])]);b.title=m.vi;b.setAttribute("aria-current",String(key(c)===sel));b.setAttribute("data-fk","row:"+key(c));
    b.onclick=function(){sel=key(c);renderRun()};panel.appendChild(b)});return panel}
function detailPanel(s){var c=s.cases.filter(function(x){return key(x)===sel})[0];var panel=h("section","panel detail");
  if(!c){panel.appendChild(h("div","none-sel",["Chọn một case để xem từng bước và ảnh bằng chứng."]));return panel}
  var head=h("div","dhead",[h("h2",0,[c.title||("Case "+c.id)]),h("div",0,[badge(c.status),document.createTextNode("  "),h("span","mono",["ID "+c.id+(c.row?" · dòng "+c.row:"")])])]);
  if(c.pre)head.appendChild(h("div","kv",[h("b",0,["Tiền điều kiện"]),c.pre]));if(c.expected)head.appendChild(h("div","kv",[h("b",0,["Kết quả mong đợi"]),c.expected]));panel.appendChild(head);
  var flat=[],pics=function(list,stepNo,cap){return list.map(function(p){flat.push({p:p,step:stepNo,cap:cap});return flat.length-1})};
  var ol=h("ol","steps");var n=Math.max(c.plannedSteps.length,c.steps.length);
  for(var i=0;i<n;i++){var stp=c.steps[i],inProgress=!stp&&c.status==="RUNNING"&&i===c.steps.length,m=stp?st(stp.status||"RUNNING"):(inProgress?st("RUNNING"):st("PLANNED"));
    var title=stp?(stp.caption==="runner error"?"Lỗi trong script test (chưa kết luận gì về ứng dụng)":stp.caption):(c.plannedSteps[i]||"");
    var body=[h("div","plan",[(i+1)+". "+(c.plannedSteps[i]||"(bước phát sinh)")])];
    if(stp){body.push(h("div","did",[title]));if(stp.actual)body.push(h("div","act",[stp.actual]));if(stp.remark)body.push(h("div","why",[stp.remark]))}
    else body.push(h("div","todo",[inProgress?"Đang làm…":"Chưa chạy"]));
    var own=stp?stp.images:(inProgress?c.liveImages:[]);
    if(own.length){var idx=pics(own,i+1,title||("Bước "+(i+1)));var th=h("div","thumbs");
      idx.forEach(function(k,j){var b=h("button");b.type="button";b.title="Xem ảnh "+(j+1)+" của bước này";var im=h("img");im.loading="lazy";im.alt="Bước "+(i+1)+", ảnh "+(j+1);im.src=imgUrl(s.id,flat[k].p);b.appendChild(im);
        b.appendChild(h("span","k num",[String(j+1)]));if(inProgress&&j===idx.length-1)b.appendChild(h("span","new",["Mới nhất"]));b.onclick=function(){openLb(s.id,flat,k)};th.appendChild(b)});body.push(th)}
    ol.appendChild(h("li","step",[icon(m.ic,m.c),h("div","sb",body)]))}
  panel.appendChild(ol);
  return panel}

/* picture viewer, inside the page: previous / next walk through every picture of the case in step order */
var lb={id:"",list:[],i:0};
function showLb(){var it=lb.list[lb.i];$("lb-img").src=imgUrl(lb.id,it.p);$("lb-open").href=imgUrl(lb.id,it.p);
  $("lb-cap").textContent="Bước "+it.step+" · "+it.cap;$("lb-count").textContent=(lb.i+1)+" / "+lb.list.length;
  var strip=$("lb-strip");Array.prototype.forEach.call(strip.children,function(b,k){b.setAttribute("aria-current",String(k===lb.i));if(k===lb.i)b.scrollIntoView({block:"nearest",inline:"center"})});
  [lb.i-1,lb.i+1].forEach(function(k){if(lb.list[k]){var pre=new Image();pre.src=imgUrl(lb.id,lb.list[k].p)}})}
function stepLb(d){lb.i=(lb.i+d+lb.list.length)%lb.list.length;showLb()}
function openLb(id,list,i){lb={id:id,list:list.slice(),i:i};var strip=$("lb-strip");strip.textContent="";
  lb.list.forEach(function(it,k){var b=h("button");b.type="button";b.title="Bước "+it.step;var im=h("img");im.alt="Bước "+it.step;im.src=imgUrl(id,it.p);b.appendChild(im);b.onclick=function(){lb.i=k;showLb()};strip.appendChild(b)});
  showLb();$("lb").classList.add("open");$("lb-close").focus()}
function closeLb(){$("lb").classList.remove("open");$("lb-img").removeAttribute("src")}
$("lb-close").onclick=closeLb;$("lb-prev").onclick=function(){stepLb(-1)};$("lb-next").onclick=function(){stepLb(1)};
$("lb").onclick=function(e){if(e.target===$("lb")||e.target.className==="stage")closeLb()};
document.addEventListener("keydown",function(e){if(!$("lb").classList.contains("open"))return;if(e.key==="Escape")closeLb();if(e.key==="ArrowLeft")stepLb(-1);if(e.key==="ArrowRight")stepLb(1)});

/* knowledge */
var KIND={login:"Đăng nhập",navigation:"Điều hướng",locator:"Chọn phần tử",data:"Dữ liệu",error:"Lỗi",fix:"Cách sửa",quirk:"Điểm lạ",rule:"Quy tắc"};
var kdata=null;
function renderKnow(){var box=$("know");box.textContent="";
  if(!kdata){box.appendChild(h("div","empty",["Đang tải…"]));return}
  var q=$("kq")?$("kq").value.toLowerCase():"",kk=$("kk")?$("kk").value:"";
  var tools=h("div","tools");var inp=h("input");inp.id="kq";inp.type="search";inp.placeholder="Tìm trong tri thức (màn hình, từ khóa, nội dung)";inp.value=q;inp.oninput=renderKnow;
  var sl=h("select");sl.id="kk";[["","Mọi loại"]].concat(Object.keys(KIND).map(function(k){return[k,KIND[k]]})).forEach(function(o){var op=h("option",0,[o[1]]);op.value=o[0];sl.appendChild(op)});sl.value=kk;sl.onchange=renderKnow;tools.appendChild(inp);tools.appendChild(sl);box.appendChild(tools);
  var rows=kdata.lines.filter(function(l){return(!kk||l.kind===kk)&&(!q||(l.screen+" "+l.kind+" "+l.tags.join(" ")+" "+l.fact).toLowerCase().indexOf(q)>=0)});
  if(!rows.length){box.appendChild(h("div","empty",[h("b",0,[kdata.lines.length?"Không có mục nào khớp":"Chưa có tri thức nào"]),kdata.lines.length?"Thử từ khóa khác.":"Sau mỗi lần test, AI ghi điều đã học (màn hình, cách chọn phần tử, quy tắc, lỗi và cách sửa) vào đây."]));return}
  var by={};rows.forEach(function(l){(by[l.screen||"Chung"]=by[l.screen||"Chung"]||[]).push(l)});
  Object.keys(by).sort().forEach(function(sc){box.appendChild(h("div","k-group",[sc]));by[sc].forEach(function(l){box.appendChild(h("div","k-item",[h("div","kind",[KIND[l.kind]||l.kind||""]),h("div",0,[l.fact,l.tags.length?h("div","tags",[l.tags.join(" · ")]):null])]))})})}

/* tabs */
function setTab(t){tab=t;$("tab-run").setAttribute("aria-selected",String(t==="run"));$("tab-know").setAttribute("aria-selected",String(t==="know"));$("view-run").hidden=t!=="run";$("view-know").hidden=t!=="know";if(t==="know")api("knowledge").then(function(d){kdata=d;renderKnow()}).catch(function(){})}
$("tab-run").onclick=function(){setTab("run")};$("tab-know").onclick=function(){setTab("know")};

/* polling: redraw only when the data changed, so what the person is reading or has open never jumps */
function refreshRuns(){return api("runs").then(function(d){lastOk=Date.now();var j=JSON.stringify(d.runs);runs=d.runs;
    if(j!==lastRuns){lastRuns=j;if(!cur&&runs.length){var s=runs.slice().sort(function(a,b){return (ts(b.started)||0)-(ts(a.started)||0)})[0];cur=s.id}renderRail()}live()}).catch(function(){live()})}
function refreshRun(){if(!cur)return Promise.resolve();return api("run","&id="+encodeURIComponent(cur)).then(function(d){lastOk=Date.now();var j=JSON.stringify(d.run);if(j===lastState&&state)return;lastState=j;state=d.run;
    if(!sel||!state.cases.some(function(c){return key(c)===sel}))sel=pickDefault(state);renderRun()}).catch(function(){})}
setTab("run");renderRail();renderRun();
refreshRuns().then(refreshRun);
setInterval(function(){refreshRuns().then(function(){if(tab==="run")refreshRun()})},1500);
})();
</script></body></html>`;
}
