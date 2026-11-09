# Plan: Excel Partner → Office Partner (Word + Excel + PowerPoint)

Trạng thái (09/11/2026): P1–P6 đã code và test trên môi trường giả lập (Linux, không có Office). `typecheck`, `build`, `npm test` (1637 test pass, 2 bỏ qua vì cần CLI thật) và `office-addin-manifest validate` đều pass. Đã thêm sau P6: cross-file giữa 3 app, đọc `.docx`/`.pptx` chưa mở, lịch sử chat theo từng app, mở lại pane = chat mới, đổi tên gói `office-partner` (xem mục 12). Đã cài lên máy Windows của chủ dự án (đường `excel-partner` → `office-partner`). **Chưa smoke test live trên Word/PowerPoint/Excel thật** và **chưa publish lên npm** — xem mục 10, 11, 12.
Nhánh: `feat/office-partner-multi-host`, đã merge vào `main-office` (tách từ `main`, chưa đẩy lên remote).
Mục tiêu: một add-in duy nhất, cài một lần, dùng được ở Word, Excel và PowerPoint; giữ nguyên hành vi Excel hiện tại; bổ sung tính năng Word/PowerPoint ngang tầm Claude in Microsoft Office.

Tài liệu này đã được đối chiếu với `@types/office-js@1.0.612` và manifest 3 host đã validate bằng `office-addin-manifest`.

---

## 0. Quyết định đã chốt

- Một manifest XML duy nhất với 3 `<Host>`: `Document`, `Workbook`, `Presentation`.
- Một codebase, một task pane, một bridge; khác nhau ở `HostAdapter`.
- Giữ nguyên toàn bộ tool và executor Excel hiện có.
- Word/PowerPoint có executor riêng, tool riêng, prompt addendum riêng.
- Không thêm dependency: dùng `fflate` (đã có) cho PPTX zip, `DOMParser`/`XMLSerializer` cho OOXML.
- Package npm đã đổi thành `office-partner` (lệnh `office-partner`, file `bin/office-partner.mjs`). Tên cũ `excel-partner` còn lại trên npm dưới dạng gói cầu nối `legacy/excel-partner/` — xem mục 12.
- File cấu hình `excel-partner.config.json` giữ nguyên tên để không phá cài đặt cũ.
- Giữ `dataDir` tên `Excel Partner` và thư mục backups hiện tại để không phá cài đặt cũ.
- Đăng ký sideload là **Office-wide** (`WEF\Developer` + trusted catalog), không cần đăng ký riêng từng app.

---

## 1. Yêu cầu tính năng

### 1.1 Bắt buộc chung (cả 3 host)

| Yêu cầu | Cách đáp ứng |
|---|---|
| Cài 1 lần, ribbon ở cả 3 app | 1 manifest, 3 `<Host>`, 3 group ribbon |
| Task pane biết host | `Office.context.host` + fallback `requirements.isSetSupported` |
| Chat với Claude Code / Copilot / Codex / OpenCode | dùng lại `bridge/`, `providers/`, `gateway/` |
| Selection chip + snapshot gửi kèm prompt | `get_selection` theo adapter từng host |
| `execute_office_js` đúng runner | `Excel.run` / `Word.run` / `PowerPoint.run` |
| Tool chỉ thuộc host đang mở | registry + filter 3 lớp (MCP, gateway, pane) |
| Backup trước khi sửa | `getFileAsync(Compressed)` → `.xlsx/.docx/.pptx` |
| Undo/restore | Excel giữ journal; Word/PPT phase 1 chỉ backup, chưa auto-restore |
| Cross-file | cả 3 app đọc được lẫn nhau (read-only) và nhờ trợ lý của file kia làm việc; xem mục 12 |
| Standalone browser | host mặc định `excel`, không crash |
| Không lộ secret / không log nội dung nhạy cảm | giữ policy hiện tại |
| Preview/diff trước thao tác phá huỷ | `propose_*` + permission gate |
| Không phụ thuộc global `toolNames` | dùng `OFFICE_HOST` trong MCP env |

### 1.2 Excel

Giữ nguyên toàn bộ tool hiện có: `get_selection`, `get_sheets`, `read_range`, `get_cell_ranges`, `search_data`, `find_errors`, `check_formula_consistency`, `write_range`, `set_cells`, `set_formula`, `find_replace`, `format_range`, `clear_range`, `insert_rows`, `modify_structure`, `manage_sheet`, `manage_names`, `page_layout`, `modify_chart`, `modify_pivot_table`, `modify_table`, shapes/connectors, `execute_office_js`, testing tools, skills, connected files.

Không đổi schema. Chỉ đổi lớp dispatch và bổ sung metadata host.

### 1.3 Word — tính năng cốt lõi

| Tool | Input chính | API / gate |
|---|---|---|
| `get_selection` | — | `document.getSelection()` — WordApi 1.1 |
| `read_doc_section` | `heading?`, `paragraph_index?`, `count?` | `body.paragraphs` — 1.1 |
| `search_doc_text` | `query`, `match_case?`, `max_results?` | `body.search` — 1.1 |
| `edit_doc_text` | `old_text`, `new_text`, `match_case?`, `occurrence?` | `range.insertText(...,"Replace")` — 1.1 |
| `edit_doc_list` | `operation`, `items[]`, `after?` | List API — 1.3, hoặc `range.insertHtml` — 1.1 |
| `collapse_blank_paragraphs` | `max` | paragraph scan ngược; `paragraph.getRange().delete()` — 1.1 |
| `propose_doc_edits` | `edits[]`, `new_text_label?`, `alternatives?` | staged, UI card review |
| `read_comments` | — | `body.getComments()` — WordApi 1.4 |
| `add_comment` | `anchor_text`, `text` | `range.insertComment()` — 1.4 |
| `reply_comment` | `comment_id`, `text` | `comment.reply()` — 1.4 |
| `resolve_comment` | `comment_id` | `comment.resolved = true` — 1.4 |
| `read_tracked_changes` | `max?` | `body.getTrackedChanges()` — 1.6; fallback `getReviewedText()` — 1.4 (chỉ text) |
| `accept_reject_changes` | `mode: "accept"\|"reject"`, `scope` | `TrackedChange.accept()/reject()` — 1.6 |
| `set_track_changes` | `enabled` | `document.changeTrackingMode` — 1.4 |
| `read_attachment_pages` | `attachment_id`, `pages[]` | PDF pipeline hiện có |

### 1.4 Word — tính năng bổ sung

| Tool | Mục đích | API / gate |
|---|---|---|
| `get_doc_outline` | headings + paragraph index cho `doc_state` | `body.paragraphs` — 1.1; `styleBuiltIn` — 1.3 |
| `read_paragraph_range` | đọc N paragraph quanh một index | 1.1 |
| `apply_style` | đổi style/format theo range | `range.styleBuiltIn` — 1.3; font — 1.1 |
| `insert_table` | tạo và điền bảng trong 1 call | `body.insertTable` — 1.1; `Table.values` — 1.5 |
| `edit_table` | thêm/xoá hàng cột, merge, format cell | add/delete rows/cols — 1.3; merge — 1.4 |
| `manage_content_controls` | liệt kê/điền placeholder template | `range.contentControls` — 1.1 |
| `insert_field` / `update_fields` | field, cross-reference | `Range.insertField` / `Field.updateResult` — 1.5 |
| `insert_toc` / `update_toc` | mục lục | `WordApiDesktop 1.4` — **desktop only** |
| `insert_hyperlink` | link trong body | `insertHtml` — 1.1 hoặc `insertField` Hyperlink — 1.5 |
| `insert_bookmark` | bookmark | `Range.insertBookmark` — 1.4 |
| `manage_headers_footers` | header/footer từng section | 1.1 |
| `insert_footnote` | footnote | `Range.insertFootnote` — 1.5 |
| `set_page_layout` | margin, orientation, columns | `Section.pageSetup` — 1.1 |
| `read_reviewed_edits` | phản hồi từ proposal card | nội bộ |

### 1.5 PowerPoint — tính năng cốt lõi

| Tool | Input chính | API / gate |
|---|---|---|
| `get_selection` | — | `getSelectedSlides()` / `getSelectedShapes()` — 1.5; fallback `getSelectedDataAsync(SlideRange)` — Selection 1.1 |
| `get_presentation_metadata` | — | `title` — 1.0; `slides` — 1.2; `pageSetup` — 1.10; fallback parse `presentation.xml` — 1.8 |
| `list_layouts` | — | `presentation.slideMasters` → `SlideMaster.layouts` — 1.3 |
| `add_slide` | `layout_id`, `position?` | `slides.add({ slideMasterId, layoutId })` — 1.3 |
| `list_slide_shapes` | `slide_id` | `slide.shapes` — 1.3 |
| `read_slide_text` | `ref` (`"258:5"`) | `shape.textFrame` — 1.4; table — 1.8 |
| `edit_slide_text` | `ref`, `text` | `shape.textFrame.textRange.text` — 1.4; giữ format đơn giản |
| `edit_slide_xml` | `slide_id`, `code` | export — 1.8 → unzip → XML → insert — 1.2 |
| `edit_slide_master` | `code` | **không hỗ trợ arbitrary XML**; chỉ sửa shape có sẵn trên master — phase 2 |
| `edit_slide_chart` | `slide_id`, `shape_id`, `code` | OOXML — phase 2, high risk |
| `insert_slide_element` | `slide_id`, spec | textbox/geometric shape/line — 1.4; table — 1.8; image — OOXML |
| `remove_slide_element` | `slide_id`, `shape_id` | `shape.delete()` — 1.3 |
| `duplicate_slide` | `slide_id` | không có `duplicate()`; export — 1.8 + insert — 1.2 |
| `move_slide` | `slide_id`, `position` | `slide.moveTo()` — 1.8 |
| `copy_image_between_slides` | `from_ref`, `to_slide_id` | không có `addImage`; OOXML/`ShapeFill.setImage` — phase 2 |
| `modify_presentation_structure` | `operation`, spec | add/delete/move; không duplicate |
| `verify_slides` | `slide_ids?` | geometry — 1.4; slide size — 1.10/fallback; contrast best effort |
| `verify_slide_visual` | `slide_id` | `getImageAsBase64` — 1.8 + fresh-context reviewer |

### 1.6 PowerPoint — tính năng bổ sung

| Tool | Mục đích | API / gate |
|---|---|---|
| `set_slide_notes` / `read_slide_notes` | speaker notes | **không có typed API**; cần OOXML spike — phase 2 |
| `set_slide_background` | nền slide | `slide.background` — 1.10; OOXML fallback |
| `set_shape_alt_text` | accessibility | `shape.altTextTitle/Description` — 1.10; OOXML fallback |
| `insert_image` | ảnh từ base64/blob | **không có `shapes.addImage`**; OOXML — phase 2 |
| `insert_slide_table` / `edit_slide_table` | bảng (tên có `slide_` để không trùng tool Word) | `addTable` — 1.8; values — 1.8; add/delete rows/cols — 1.9 |
| `set_slide_hyperlink` | link trên shape/text (tên không trùng `insert_hyperlink` của Word) | `TextRange.setHyperlink` — 1.10 |
| `set_shape_zorder` | thứ tự lớp | `setZOrder` — 1.8 |
| `group_shapes` / `ungroup_shapes` | nhóm | `Shape.group`/`ungroup` — 1.8 |
| `align_shapes` | căn hàng/cột và chia đều khoảng cách (tham số `align` hoặc `distribute`, không có tool `distribute_shapes` riêng) | tính `left/top` — 1.4 |
| `set_slide_transition` | chuyển slide | OOXML — phase 2 |
| `manage_sections` | section của deck | **không khả thi qua Office.js**; cần whole-file replace |
| `export_slide_image` | ảnh slide | `getImageAsBase64` — 1.8 |
| `insert_icon` / `search_icons` | icon vector | icon library — phase 2 |

### 1.7 Tool nguồn gốc

| Nhóm | Nguồn |
|---|---|
| Excel: `get_sheets_metadata`, `get_cell_ranges`, `get_range_as_csv`, `set_cell_range`, `clear_cell_range`, `modify_sheet_structure`, `execute_office_js` | Claude (bundle + prompt) |
| Excel còn lại | repo hiện tại |
| Word: `edit_doc_text`, `edit_doc_list`, `collapse_blank_paragraphs`, `propose_doc_edits`, `read_doc_section`, `search_doc_text`, `read_attachment_pages`, `execute_office_js` | Claude (bundle + prompt) |
| Word: comments, track changes, outline, fields, TOC, headers/footers, content controls | tự thiết kế (Claude làm bằng `execute_office_js`) |
| PPT: `list_slide_shapes`, `read_slide_text`, `edit_slide_text`, `edit_slide_xml`, `insert_slide_element`, `remove_slide_element`, `duplicate_slide`, `move_slide`, `modify_presentation_structure`, `verify_slides`, `verify_slide_visual`, `execute_office_js` | Claude (bundle + prompt) |
| PPT: notes, background, alt text, image, transition, sections, export image | tự thiết kế (Claude làm bằng `execute_office_js`/OOXML) |
| PPT: `insert_slide_table` / `edit_slide_table` | typed table API 1.8/1.9; không dùng OOXML cho cell |
| `store_blob`, `container_upload`, `code_execution` | Claude nội bộ — không đưa vào |

---

## 2. Kiến trúc

```
Office host (Word / Excel / PowerPoint)
        │  Office.js
┌───────▼─────────────────────────────────────┐
│ addin/taskpane                              │
│  host/detect.ts      → "excel|word|ppt"     │
│  host/types.ts       → HostAdapter          │
│  host/excel.ts       → executor hiện tại    │
│  host/word.ts        → executor/word/*      │
│  host/powerpoint.ts  → executor/powerpoint/* │
│  main.ts             → UI dùng chung        │
└───────┬─────────────────────────────────────┘
        │ WebSocket + clientInfo.host
┌───────▼─────────────────────────────────────┐
│ bridge (một process)                        │
│  WorkbookSession.host                       │
│  ToolGateway: validate/gate theo host        │
│  mcpEndpointFor → OFFICE_HOST                │
│  system prompt + selection injection         │
└───────┬─────────────────────────────────────┘
        │ stdio MCP
┌───────▼─────────────────────────────────────┐
│ excel_mcp (mỗi provider session)             │
│  toolsForHost(OFFICE_HOST)                  │
│  ListTools chỉ trả tool của host            │
└─────────────────────────────────────────────┘
```

Không đặt `toolNames` global trong `McpEndpoint`: việc này chặn `AskUserQuestion` của Claude và đổi hành vi `clientTools` của OpenCode. Host filter phải nằm ở `OFFICE_HOST` + `ListTools`.

### File map

```
protocol/src/
  host.ts                 # HostId, TOOL_HOSTS, toolsForHost()
  tools-word.ts           # zod schemas Word
  tools-ppt.ts            # zod schemas PPT
  registry.ts             # gộp 3 nguồn, export theo host
  tool-results.ts         # selection snapshot mở rộng
  messages.ts             # ClientInfo.host

addin/taskpane/
  host/detect.ts
  host/types.ts
  host/excel.ts
  host/word.ts
  host/powerpoint.ts
  executor/word/*.ts
  executor/powerpoint/*.ts
  executor/index.ts       # dispatch theo host
  executor/backup.ts      # getFileAsync tổng quát
  main.ts                 # bỏ import trực tiếp Excel

excel_mcp/src/
  tool-definitions-word.ts
  tool-definitions-ppt.ts
  server.ts               # filter theo OFFICE_HOST
  cli.ts                  # đọc OFFICE_HOST

bridge/src/
  session.ts              # host
  tool-gateway.ts         # validate/gate theo host
  server.ts               # mcpEndpointFor + prompt injection
  document-backups.ts     # thay workbook-backups (file cũ còn lại làm alias), key theo ext
  selection.ts            # selection theo host

providers/src/
  host-prompts.ts         # addendum prompt theo host (excel/word/powerpoint)

install.ps1 / install.sh / bin/office-partner.mjs
  # đăng ký 3 host, tên file backup theo host
```

---

## 3. Công thức triển khai (recipe bắt buộc)

### C1. Host detection

```ts
export type HostId = "excel" | "word" | "powerpoint";

export function detectHost(): HostId {
  const h = Office.context.host as string | undefined;
  if (h === "Excel") return "excel";
  if (h === "Word") return "word";
  if (h === "PowerPoint") return "powerpoint";
  const req = Office.context.requirements;
  if (req.isSetSupported("ExcelApi", "1.1")) return "excel";
  if (req.isSetSupported("WordApi", "1.1")) return "word";
  if (req.isSetSupported("PowerPointApi", "1.1")) return "powerpoint";
  return "excel"; // standalone / browser preview
}
```

Không dùng URL hay đuôi file làm nguồn chính; chỉ fallback khi `Office.context.host` undefined.

### C2. Manifest đa host

Manifest hiện tại: `addin/manifest.xml`, ID `7f2b6c1e-9a4d-4c3f-8b21-2d5e6f7a8b90`, version `1.3.0.0`, chỉ `Workbook` + `ExcelApi 1.9`. Giữ nguyên `Id`, tăng `Version` lên `1.4.0.0`.

Base requirements chỉ dùng common requirement set có mặt ở cả 3 host:

```xml
<Hosts>
  <Host Name="Document"/>
  <Host Name="Workbook"/>
  <Host Name="Presentation"/>
</Hosts>
<Requirements>
  <Sets DefaultMinVersion="1.1">
    <Set Name="File" MinVersion="1.1"/>
    <Set Name="DocumentEvents" MinVersion="1.1"/>
    <Set Name="Selection" MinVersion="1.1"/>
  </Sets>
</Requirements>
```

VersionOverrides:

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
                  xsi:type="VersionOverridesV1_0">
  <Hosts>
    <Host xsi:type="Document">
      <DesktopFormFactor>
        <!-- ribbon Word: group OP.Word.Group, button OP.Word.OpenPane, TaskpaneId OP.Word.Pane -->
      </DesktopFormFactor>
    </Host>
    <Host xsi:type="Workbook">
      <DesktopFormFactor>
        <!-- ribbon Excel: group OP.Excel.Group, button OP.Excel.OpenPane, TaskpaneId OP.Excel.Pane -->
      </DesktopFormFactor>
    </Host>
    <Host xsi:type="Presentation">
      <DesktopFormFactor>
        <!-- ribbon PowerPoint: group OP.Ppt.Group, button OP.Ppt.OpenPane, TaskpaneId OP.Ppt.Pane -->
      </DesktopFormFactor>
    </Host>
  </Hosts>
  <Resources>
    <!-- icon, URL, label dùng chung một lần -->
  </Resources>
</VersionOverrides>
```

Quy tắc:

- Không liệt kê `ExcelApi + WordApi + PowerPointApi` cùng lúc trong base `<Requirements>`.
- **`VersionOverrides` không có `<Requirements>`**; requirements chỉ nằm ở base. Phần còn lại gate bằng runtime.
- Mỗi host có `Group id`, `Button id`, `TaskpaneId` riêng.
- `<Resources>` dùng chung để tránh lặp icon/label.
- Validator có thể báo min version thấp vì base requirements là common set; UI phải tự chặn host thiếu API.

### C3. Runtime requirement gate

```ts
const FEATURES = {
  excel:      [["ExcelApi", "1.9"]],
  word:       [["WordApi", "1.1"]],
  powerpoint: [["PowerPointApi", "1.2"]],
} as const;

const WORD_COMMENTS = [["WordApi", "1.4"]];
const WORD_FIELDS   = [["WordApi", "1.5"]];
const WORD_TRACK    = [["WordApi", "1.6"]];
const WORD_BOOKMARK = [["WordApi", "1.4"]];
const WORD_TOC      = [["WordApiDesktop", "1.4"]];

const PPT_SELECTION = [["PowerPointApi", "1.5"]];
const PPT_EXPORT    = [["PowerPointApi", "1.8"]];
const PPT_MODERN    = [["PowerPointApi", "1.10"]];

function available(feature: readonly (readonly [string, string])[]): boolean {
  return feature.every(([set, version]) =>
    Office.context.requirements.isSetSupported(set, version)
  );
}
```

Không gọi API vượt min version mà không check. Thiếu requirement → trả `unsupported` kèm hướng dẫn, không crash.
TOC là `WordApiDesktop 1.4` nên **desktop-only**; Word web phải trả unsupported.

### C4. Tool registry + exposure

```ts
// protocol/src/host.ts
export const TOOL_HOSTS: Record<string, readonly HostId[]> = {
  ...Object.fromEntries(EXCEL_TOOLS.map(name => [name, ["excel"]])),
  ...Object.fromEntries(WORD_TOOLS.map(name => [name, ["word"]])),
  ...Object.fromEntries(PPT_TOOLS.map(name => [name, ["powerpoint"]])),
  get_selection:     ["excel", "word", "powerpoint"],
  execute_office_js: ["excel", "word", "powerpoint"],
};

export function toolsForHost(host: HostId, opts: {
  allowOptIn: boolean;
  webSearch?: boolean;
  playwright?: boolean;
}): string[] {
  return Object.keys(TOOL_HOSTS).filter(name =>
    TOOL_HOSTS[name].includes(host) &&
    (opts.allowOptIn || name !== "execute_office_js") &&
    (opts.playwright || !BROWSER_TOOLS.includes(name)) &&
    (host === "excel" || !EXCEL_ONLY_SERVER_TOOLS.includes(name))
  );
}
```

- `create_artifact` do pane xử lý, **không nằm trong `SERVER_TOOLS`**; vẫn expose cho Word/PPT.
- Tool chỉ Excel gồm `read_range_image`, `insert_sheets_from_file` và các tool đọc/ghi workbook. `list_connected_files`, `read_connected_file`, `send_to_connected_file`, `read_workbook_file` và `read_document_file` là host-agnostic (cả 3 app).
- Skills, Python, Playwright, `read_attachment_pages`, `create_artifact` là host-agnostic.
- `exposedToolNamesForHost(host, opts)` là nguồn duy nhất cho MCP `ListTools` và provider fallback.

### C5. End-to-end dispatch

```
1. Office.onReady → detectHost()
2. hello { clientInfo: { host } }
3. Session.host = clientInfo.host
4. mcpEndpointFor(session) → env { OFFICE_HOST: session.host, ... }
5. MCP ListTools = toolsForHost(OFFICE_HOST)
6. Model gọi tool → MCP validate zod → bridge /tool
7. ToolGateway: isToolNameForHost(host, tool) → parseToolInput(host, tool)
8. session.requestTool → WebSocket tool_request
9. pane executeTool(tool, args) → adapter[host] → tool_result
```

`mcpEndpointFor` không set `toolNames`. Provider fallback (`copilot`, `copilot-native`) phải gọi `exposedToolNamesForHost`; Claude/OpenCode để MCP discovery tự lọc.

### C6. Selection snapshot

Giữ `SelectionSnapshotSchema` tương thích ngược, thêm `host` và các field optional:

```ts
const ExcelSelectionSnapshotSchema = SelectionSnapshotSchema.extend({
  host: z.literal("excel").default("excel"),
});

const WordSelectionSnapshotSchema = z.object({
  host: z.literal("word"),
  text: z.string(),
  cursor: z.boolean(),
  style: z.string().optional(),
  paragraphs: z.number().int(),
  highlight: z.string().nullable().optional(),
  comments: z.array(z.object({
    id: z.string(),
    text: z.string(),
    replies: z.number().int(),
    resolved: z.boolean(),
  })).optional(),
  trackChanges: z.boolean().optional(),
  truncated: z.boolean(),
});

const PptSelectionSnapshotSchema = z.object({
  host: z.literal("powerpoint"),
  slides: z.array(z.object({
    id: z.string(),
    position: z.number().int(),
    title: z.string(),
  })),
  shapes: z.array(z.object({
    ref: z.string(),
    id: z.string(),
    name: z.string(),
    type: z.string(),
  })),
  truncated: z.boolean(),
});

export const AnySelectionSnapshotSchema = z.preprocess(
  (value) => {
    if (value && typeof value === "object" && !("host" in value)) {
      return { ...value, host: "excel" };
    }
    return value;
  },
  z.discriminatedUnion("host", [
    ExcelSelectionSnapshotSchema,
    WordSelectionSnapshotSchema,
    PptSelectionSnapshotSchema,
  ]),
);
```

Prompt injection (`providers/src/prompt.ts`):

```xml
<selection_snapshot>{json}</selection_snapshot>
```

- Tag hiện tại là `<selection_snapshot>`, **không phải** `<user_selection>`; test T1 đang phụ thuộc tag này.
- `buildTurnPrompt` phải branch theo `host`; chỉ gọi `selection.values.flat()` khi `host === "excel"`.
- Bridge chỉ chạy `enrichWindowsSelection` khi `host === "excel"`.
- PPT event `DocumentSelectionChanged` có fire, nhưng shape selection có thể không đổi event; dùng debounce + focus + before-send fallback.
- PPT < 1.5: chỉ lấy selected slide qua `getSelectedDataAsync(SlideRange)`, không có shape selection.

### C7. PPT OOXML round-trip

```ts
const before = await ctx.presentation.slides.load("items/id");
await ctx.sync();
const beforeIds = new Set(before.items.map(s => s.id));

// 1. export slide hiện tại thành pptx một slide
const b64 = await slide.exportAsBase64();          // PowerPointApi 1.8
// 2. unzip bằng fflate
const files = unzipSync(base64ToBytes(b64));       // Record<string, Uint8Array>
const xml = strFromU8(files["ppt/slides/slide1.xml"]);
// 3. sửa bằng DOM, không string concat
const doc = new DOMParser().parseFromString(xml, "application/xml");
mutate(doc);                                        // namespace-aware
files["ppt/slides/slide1.xml"] =
  strToU8(new XMLSerializer().serializeToString(doc));
// 4. re-zip
const out = bytesToBase64(zipSync(files));
// 5. insert; API trả void, không trả collection
await ctx.presentation.insertSlidesFromBase64(out, {
  targetSlideId: previousSlideId, // bỏ trống nếu chèn ở đầu
});
await ctx.sync();

// 6. tìm slide mới bằng ID diff
const after = await ctx.presentation.slides.load("items/id");
await ctx.sync();
const inserted = after.items.filter(s => !beforeIds.has(s.id));
```

Quy tắc:

- `insertSlidesFromBase64` trả **void**; không có `inserted` collection và không có `resolveSlideId`.
- Muốn xoá slide cũ: xoá sau khi insert và sync, dùng ID cũ đã lưu.
- `targetSlideId` là slide ngay trước vị trí cũ; nếu là slide đầu, bỏ `targetSlideId` để chèn ở đầu.
- Kiểm tra external relationship trước khi insert (`IllegalAccessError`).
- Giữ đúng thứ tự child trong OOXML; dùng `insertBefore`, không `appendChild` tuỳ tiện.
- Không dùng recipe này cho `edit_slide_master`; arbitrary master XML không khả thi.
- Luôn backup trước khi round-trip.

### C8. Word edit + tracked changes

```ts
await Word.run(async context => {
  if (track && available(WORD_TRACK)) {
    context.document.changeTrackingMode =
      Word.ChangeTrackingMode.trackAll;
  }

  const hits = context.document.body.search(oldText, { matchCase });
  hits.load("items");
  await context.sync();

  if (hits.items.length === 0) {
    throw new ToolFailure("not_found", "old_text does not appear in the document");
  }

  const target = occurrence ? hits.items[occurrence] : hits.items[0];
  target.insertText(newText, Word.InsertLocation.replace);
  await context.sync();
});
```

Quy tắc:

- Không dùng `body.clear()` / rebuild cho một chỉnh sửa text.
- Bulk replace: thêm `replace_all_doc_text` hoặc `edit_doc_text({ all: true })`; không loop `body.search` trong một `Word.run`.
- `propose_doc_edits` chỉ stage; user bấm Apply mới ghi.
- Comment: reply/resolve **trước**, edit anchored text **sau** (edit làm mất comment).
- `Paragraph` không có `delete()`; dùng `paragraph.getRange().delete()`.
- List: dùng List API 1.3 hoặc `insertHtml` 1.1; không giả định `insertParagraph` + style là đủ numbering.
- Hyperlink: không có `Range.insertHyperlink`; dùng `insertHtml` hoặc field `Hyperlink`.
- TOC: `WordApiDesktop 1.4`, desktop-only; không có `Range.insertTableOfContents`.
- Tracked changes: đọc/accept/reject cần `WordApi 1.6`; `getReviewedText` 1.4 chỉ là fallback text.

### C9. Backup + restore

```ts
Office.context.document.getFileAsync(
  Office.FileType.Compressed,
  { sliceSize: 4 * 1024 * 1024 },
  callback,
);
```

- `File` + `CompressedFile` dùng được cho `.xlsx/.docx/.pptx`, nhưng Word/PowerPoint web cần test và phải bắt lỗi.
- Extension: `{ excel: ".xlsx", word: ".docx", powerpoint: ".pptx" }`.
- Giới hạn backup: 100 MB; slice 4 MB.
- Phase 1: **chỉ backup**, không auto-restore cho Word/PPT.
- Backup một lần trước write đầu tiên của mỗi reply; Word/PPT dùng `UndoJournal` ở chế độ backup-only, không auto-restore.
- Excel: giữ `UndoJournal` hiện tại; `journal.ts`/`undo-plan.ts` không áp dụng nguyên trạng cho Word/PPT.
- Word restore (phase 2, nếu có bridge read-back): `body.clear()` → `body.insertFileFromBase64(base64, "End")`; giới hạn 4 MB, mất comments/tracked changes/settings.
- PPT restore (phase 2): xoá toàn bộ slide → `insertSlidesFromBase64(backup)`; mất section/master tùy file.
- Bridge chưa có endpoint đọc backup bytes; muốn restore tự động phải thêm endpoint ở phase 2.

### C10. Prompt/context theo host

- Giữ `SYSTEM_PROMPT` người dùng chỉnh.
- Thêm addendum theo host từ `providers/src/host-prompts.ts`, nối sau prompt người dùng.
- Word: selection là "pointer", track-changes granularity, không rebuild section.
- PowerPoint: edit-in-place, shape ref, verify 2 lớp, không xoá slide để rebuild.
- Excel: giữ nguyên `RECOMMENDED_SYSTEM_PROMPT` hiện tại.
- Context tag duy nhất hiện có là `<selection_snapshot>`. `<user_changes>`, `<user_reviewed_edits>` và `<doc_state>` **chưa tồn tại trong code**; chỉ thêm khi thực sự implement.
- Văn bản lấy từ file (ô, đoạn Word, chữ trên slide) nằm trong JSON của snapshot. `providers/src/prompt.ts` viết `<` thành `\u003c` để đoạn như `</selection_snapshot>` trong file không đóng được khối; có test ở `tests/host.test.ts`.

### C11. Ghép nối code hiện tại (bắt buộc)

| Khu vực | Hiện trạng | Việc phải sửa |
|---|---|---|
| `addin/manifest.xml` | Excel-only, ID `7f2b6c1e-9a4d-4c3f-8b21-2d5e6f7a8b90`, v`1.3.0.0`, `ExcelApi 1.9` | 3 hosts, v`1.4.0.0`, base common sets, per-host IDs |
| `scripts/sideload.ps1` | Ghi `WEF\Developer` + trusted catalog, đăng ký Office-wide | Giữ nguyên; không đăng ký riêng từng app; chỉ đổi text/branding |
| `install.ps1` / `scripts/install.ps1` / `bin/office-partner.mjs` | Text và mở Excel; backups CLI `.xlsx` | Thêm Word/PowerPoint optional; backups CLI nhận `.xlsx/.docx/.pptx`; giữ `dataDir` |
| `protocol/src/messages.ts` | `selection` là `SelectionSnapshotSchema.nullable()` (quanh dòng 83) | Đổi sang `AnySelectionSnapshotSchema.nullable()`; `ClientInfo.host` default `"excel"` |
| `protocol/src/registry.ts` | Registry Excel phẳng | Thêm `TOOL_HOSTS`, `toolsForHost`, `exposedToolNamesForHost`; giữ export cũ |
| `bridge/src/tool-gateway.ts` | Validate Excel-only; `get_selection` parse schema Excel (quanh dòng 210) | Host-aware; sai host trả `unknown_tool`/`unsupported` |
| `bridge/src/session.ts` | `WorkbookSession` chưa có host | Thêm `host`; standalone default `"excel"` |
| `bridge/src/server.ts` | Selection enrichment chạy mọi host (quanh dòng 1221); message Excel-centric | Chỉ enrich COM khi `host === "excel"`; `mcpEndpointFor` thêm `OFFICE_HOST`; **không** set `toolNames`; `DISCONNECTED_MID_CALL` generic/host-aware; giữ `workbook_info` naming |
| `bridge/src/workbook-backups.ts` | Ép `.xlsx`, chỉ nhận zip/PK, strip Excel ext | Đổi thành `document-backups.ts`; hỗ trợ `.xlsx/.docx/.pptx`; lỗi generic; CLI cập nhật |
| `bridge/src/windows-selection.ts` | Excel COM | Giữ Excel-only |
| `providers/src/prompt.ts` | `selection.values.flat()` vô điều kiện (quanh dòng 47) | Branch theo `host`; tag `<selection_snapshot>` |
| `providers/src/claude.ts` | `toolNames` chặn `AskUserQuestion` (quanh dòng 218) | Không set global; host filter qua `OFFICE_HOST` |
| `providers/src/opencode.ts` | `clientTools` phụ thuộc `toolNames` (quanh dòng 214) | Không set global; dùng registry host-aware |
| `providers/src/copilot.ts`, `copilot-native.ts` | Fallback dùng danh sách Excel | Dùng `exposedToolNamesForHost` |
| `excel_mcp/src/server.ts` | `ALL_TOOL_DEFINITIONS` phẳng | Thêm host; combine + filter; test đồng bộ protocol |
| `addin/taskpane/main.ts` | Gọi trực tiếp Excel: `currentWorkbookName`, `refreshSheetNames`, `selectRange`, `goTo`, `watchSelection`, `watchShapeSelectionEvents`, `readWorkbookName`, `buildContextPreface` | `HostAdapter`; `executeTool(host, ...)`; `WRITE_TOOLS` union; `getSelectionLabel` |
| `addin/taskpane/executor/office-js.ts` | Hard-code `AsyncFunction("context","Excel","Office",code)` và `Excel.run` | Thêm `Word.run`/`PowerPoint.run` + global tương ứng |
| `addin/taskpane/executor/journal.ts`, `undo-plan.ts` | Excel-only | Word/PPT dùng lại `UndoJournal` ở chế độ backup-only; một lần mỗi reply |
| `addin/taskpane/executor/backup.ts` | `readWorkbookFile()` qua `getFileAsync(Compressed)` | Generic tên hàm; extension theo host; 100 MB; bắt lỗi web |
| `gateway/src/engine.ts`, `relay.ts` | Gateway độc lập | Không cần host filter |
| `bin/office-partner.mjs` | CLI backups `.xlsx` | Liệt kê 3 ext; chưa có restore-from-backup endpoint |

---

## 4. Checklist bắt buộc (Definition of Done)

Ký hiệu: `[x]` đã kiểm bằng lệnh hoặc test tự động trong repo; `[~]` đã có test với Office giả, **còn chờ Office thật** (mục 11); `[ ]` chưa kiểm.

### 4.1 Manifest và cài đặt

- [x] `office-addin-manifest validate addin/manifest.xml` pass.
- [ ] Sideload Windows: ribbon xuất hiện ở Word, Excel, PowerPoint.
- [~] Mở task pane ở cả 3, `hello` gửi `host` đúng.
- [x] Base `<Requirements>` không chứa `ExcelApi|WordApi|PowerPointApi`.
- [x] `VersionOverrides` không có `<Requirements>`.
- [x] Mỗi host có `Group id` + `TaskpaneId` riêng; `<Resources>` dùng chung.
- [x] `Id` giữ nguyên, `Version` tăng (1.4.0.0).
- [~] HTTPS + `AppDomains` + icon load đúng ở cả 3.
- [ ] Cài trên máy chưa có add-in cũ không bị trùng ribbon.
- [ ] Kiểm tra registry/catalog không tạo đăng ký trùng.

### 4.2 Tool isolation

- [x] Word không thấy `read_range`, `write_range`, `set_formula`…
- [x] Excel không thấy `read_doc_section`, `edit_slide_xml`…
- [x] PowerPoint không thấy tool Word/Excel.
- [x] Gọi tool sai host qua MCP → `unknown_tool`, không crash.
- [x] `get_selection` + `execute_office_js` có ở cả 3.
- [x] `execute_office_js` chạy đúng runner.
- [x] MCP `ListTools` khớp với bridge gate.
- [x] Không set `toolNames` global; Claude vẫn thấy `AskUserQuestion`.

### 4.3 Selection

- [x] Excel: range → chip + snapshot đúng; shape/chart → `InvalidSelection` xử lý như cũ.
- [~] Word: bôi đen text → snapshot có paragraph/style/highlight; cursor-only → cursor.
- [~] PowerPoint: chọn slide/shape → snapshot có slide id + shape ref; PPT < 1.5 fallback active slide.
- [~] Đổi selection → chip cập nhật event-driven; PPT có focus/before-send fallback.
- [x] Selection gửi kèm prompt dưới `<selection_snapshot>` đúng host.
- [x] Bridge không enrich COM cho Word/PowerPoint.
- [x] Legacy snapshot thiếu `host` được preprocess thành Excel.

### 4.4 Chức năng

- [x] Excel regression: toàn bộ test hiện tại pass.
- [~] Word: đọc section, search, edit text, list, track changes, comments.
- [~] Word: `propose_doc_edits` hiện card, chưa ghi khi chưa Apply.
- [~] Word: content controls, fields, headers/footers chạy; TOC chỉ desktop.
- [~] PowerPoint: list/read/edit text; edit text đơn giản giữ format.
- [~] PowerPoint: `edit_slide_xml` round-trip không hỏng file, verify lại.
- [~] PowerPoint: duplicate/move/delete slide đúng; duplicate dùng export+insert.
- [~] PowerPoint: `verify_slides` phát hiện overlap/overflow.
- [~] PowerPoint: table typed API chạy; notes/image/icon phase 2.
- [~] Backup tạo đúng `.xlsx/.docx/.pptx`; restore Excel chạy, Word/PPT ghi rõ limitation.
- [x] Save As/rename: ghi rõ không phát hiện đáng tin cậy; hiển thị tên từ URL.
- [x] Cross-file chạy giữa Word, Excel và PowerPoint (`tests/cross-host-files.test.ts`); tool kiểm thử bằng trình duyệt vẫn chỉ hiện ở Excel.
- [x] Standalone browser không crash.

### 4.5 An toàn và chất lượng

- [x] Nội dung cell/doc/slide coi là data, có cảnh báo prompt injection.
- [x] `execute_office_js` vẫn bị gate bởi `BRIDGE_ALLOW_OFFICE_JS`.
- [ ] Không log token/secret/nội dung file.
- [ ] Tool write chỉ chạy khi user yêu cầu.
- [~] Destructive op có permission gate / preview.
- [x] Giới hạn kết quả 30k chars áp dụng cho Word/PowerPoint.
- [x] Văn bản không tin cậy được đánh dấu rõ trong prompt.

### 4.6 Build/CI

- [x] `npm run typecheck` pass (2 tsconfig).
- [x] `npm run build` pass (server + addin + gateway).
- [x] `npm test` pass.
- [x] Test mới: host mapping, registry filter, manifest parse, Word/PPT schema.
- [ ] Live smoke: 3 host, ít nhất 1 read + 1 write mỗi host.

---

## 5. Phases và acceptance

| Phase | Nội dung | Acceptance |
|---|---|---|
| P0 | Spike OOXML: notes, image, chart, master; đo `exportAsBase64` và layout duplication | Kết luận rõ feasible/không; cập nhật plan |
| P1 | `protocol/src/host.ts`, registry 3 nguồn, `ClientInfo.host`, session.host, `OFFICE_HOST` | MCP trả đúng tool list theo env; unit test pass |
| P2 | Manifest 3 host + runtime gate + task pane detect + dispatch | Cài và mở pane ở cả 3; tool isolation đúng |
| P3 | Word executor cốt lõi + selection + prompt addendum | 10 tool Word chạy trên file thật; track changes đúng |
| P4 | PowerPoint executor cốt lõi + selection + OOXML round-trip + verify | Sửa text/XML/slide không hỏng file; selection đúng |
| P5 | Backup/undo theo host, cross-file gate, branding/installer | Backup 3 host; Excel restore; cài 1 lần hiện 3 ribbon |
| P6 | Tính năng bổ sung Word/PPT + hardening + test matrix | Checklist mục 4 pass, trừ mục phase 2 |

---

## 6. Test matrix

| Loại | Nội dung |
|---|---|
| Unit | `detectHost`, `toolsForHost`, `isToolNameForHost`, schema parse, manifest XML parse |
| Contract | MCP `ListTools` ⊆ registry của host; gọi sai host → lỗi |
| Integration (mock Office) | dispatch đúng runner; selection mapping; backup extension |
| Live Excel | regression toàn bộ tool hiện có |
| Live Word | read/search/edit/list/track changes/comments/propose |
| Live PowerPoint | list/read/edit text/XML/duplicate/move/verify |
| Negative | PPT < 1.8, Word < 1.4/1.6, TOC trên web → `unsupported` có hướng dẫn |
| Negative | Word restore > 4 MB → lỗi rõ; PPT `getFileAsync` fail trên web |
| Security | prompt injection trong cell/doc/slide; destructive tool chưa approve |

---

## 7. Giới hạn đã biết

- PowerPoint: `exportAsBase64` / `getImageAsBase64` cần 1.8+; `getSelectedSlides` cần 1.5+; `pageSetup`/`background`/`altText`/`hyperlinks` mới cần 1.10.
- PowerPoint: `Slide.delete` 1.2; `Slide.index`/`moveTo` 1.8; `presentation.properties` 1.7; `presentation.id` 1.5; `SlideLayout.type` 1.8; `SlideLayout.background` 1.10.
- PowerPoint: `getSelectedTextRange`/`setSelectedSlides` 1.5; `exportAsBase64Presentation` (whole deck) 1.10.
- PowerPoint: không có `Slide.duplicate()`; duplicate = export + insert; ID slide mới không biết trước, phải diff.
- PowerPoint: `insertSlidesFromBase64` trả void; không có `resolveSlideId`.
- PowerPoint: `exportAsBase64` là single-slide; chưa rõ có duplicate layout/master/media khi reinsert — cần spike.
- PowerPoint: không có typed notes API; notes phụ thuộc OOXML và chưa chắc `exportAsBase64` giữ `notesSlide`.
- PowerPoint: `edit_slide_master` arbitrary XML và `manage_sections` không khả thi qua Office.js.
- PowerPoint: không có `shapes.addImage`/`addSvg`; image/icon cần OOXML — phase 2.
- PowerPoint: table typed API có sẵn (1.8/1.9), không cần OOXML cho cell.
- PowerPoint: `TextRange.text` có thể làm mất format run-level; edit format-preserving cần OOXML.
- Word: TOC là `WordApiDesktop 1.4`, desktop-only; không có `Range.insertTableOfContents`.
- Word: không có `Range.insertHyperlink`; `getHyperlinkRanges` 1.3; dùng `insertHtml` 1.1 hoặc field Hyperlink 1.5.
- Word: `insertBookmark` 1.4 (không phải 1.1); `Document.acceptAllRevisions` desktop 1.4.
- Word: `insertFileFromBase64` giới hạn 4 MB, mất settings/comments/tracked changes, fail với ActiveX/form fields.
- Word: restore backup không đáng tin cậy; phase 1 chỉ backup.
- Word: `Paragraph` không có `delete()`; dùng `paragraph.getRange().delete()`.
- Word: vòng lặp `body.search` trong một `Word.run` treo với file lớn → phải dùng tool batch.
- Word/PowerPoint: không có API tên file đáng tin cậy; `FileProperties.url` chỉ có URL; tên lấy từ basename URL, unsaved dùng fallback; Save As/rename không phát hiện chắc chắn.
- Cross-file giữa các app chỉ **đọc** (và nhờ trợ lý của file kia làm); không có tool ghi xuyên file. Link ô (`cell:`) chỉ dành cho workbook.
- Icon library cần bộ asset riêng → phase 2.
- `getFileAsync(Compressed)` trên Word/PowerPoint web cần kiểm chứng thực tế.

---

## 8. Rủi ro và giảm thiểu

| Rủi ro | Giảm thiểu |
|---|---|
| Validator từ chối manifest 3 host | Chỉ common set ở base; runtime gate; validate trước commit |
| Tool Excel lọt sang Word | Filter 3 lớp: MCP ListTools, gateway, pane dispatch |
| Session lẫn host khi mở nhiều app | `host` gắn vào session lúc `hello`; mọi gate đọc từ session |
| `toolNames` global chặn `AskUserQuestion` | Không set global; dùng `OFFICE_HOST` + MCP ListTools |
| PPT round-trip hỏng file | Backup trước; DOMParser; ID diff; verify + rollback từ backup |
| Prompt Excel cũ làm model gọi tool sai | Host addendum + tool list đã lọc; test live 3 host |
| Undo Word/PowerPoint không hoàn hảo | Ghi rõ limitation; luôn backup trước write; phase 2 restore |
| Provider CLI không thấy tool mới | Kiểm tra `ListTools` qua MCP; thêm test contract |
| Khác biệt Office version giữa 3 app | Runtime gate theo từng feature; fallback rõ ràng |
| Backup không đọc lại được từ bridge | Phase 1 manual restore; phase 2 thêm endpoint |

---

## 9. Lệnh kiểm tra

```bash
npm run typecheck
npm run build
npm test
npx office-addin-manifest validate addin/manifest.xml
```

---

## 10. Trạng thái triển khai (cập nhật sau khi code)

### Đã làm

- P1: `protocol/src/host.ts` (HostId, TOOL_HOSTS, toolsForHost, exposedToolNamesForHost, isToolNameForHost, parseToolInputForHost, hostFromEnv), `ClientInfo.host`, `AnySelectionSnapshotSchema`, session.host, gate theo host trong `ToolGateway`, `OFFICE_HOST` trong MCP env, `excel_mcp` lọc `ListTools` theo host, provider fallback (`copilot`, `copilot-native`). Không set `toolNames` global.
- P2: manifest 3 host (Id giữ nguyên, v1.4.0.0, base requirements là common set, group/button/pane riêng từng host, `office-addin-manifest validate` pass); `addin/taskpane/host/*` (detect, requirements, HostAdapter cho 3 host); `bridge-client` gửi `clientInfo.host`; `execute_office_js` chạy `Excel.run` / `Word.run` / `PowerPoint.run`.
- P3 (Word, 27 tool): `protocol/src/tools-word.ts`, `addin/taskpane/executor/word/*`, `excel_mcp/src/tool-definitions-word.ts`, card duyệt `propose_doc_edits` (`ui/proposal.ts`).
- P4 (PowerPoint, 24 tool): `protocol/src/tools-ppt.ts`, `addin/taskpane/executor/powerpoint/*` (gồm OOXML round-trip bằng fflate + DOMParser), `excel_mcp/src/tool-definitions-ppt.ts`.
- P5: `bridge/src/document-backups.ts` (`.xlsx/.docx/.pptx`, `workbook-backups.ts` giữ lại làm alias), backup-only cho Word/PPT bằng `UndoJournal` hiện có (mọi write tool = "không ghi lại", một bản sao mỗi reply), cross-file chỉ Excel, host prompt (`providers/src/host-prompts.ts`), cài đặt/branding.
- Test mới: `host*.test.ts`, `manifest.test.ts`, `document-backups.test.ts`, `host-prompts.test.ts`, `pane-boot.test.ts` (pane khởi động ở 3 host với Office giả), `mcp-host.stdio.test.ts` (tiến trình MCP thật theo `OFFICE_HOST`), `word-*.test.ts`, `ppt-*.test.ts`.

### Khác với plan ban đầu

- `edit_slide_xml` chạy code của model nên thuộc nhóm opt-in cùng `execute_office_js` (`OPT_IN_TOOLS`, tắt bằng `BRIDGE_ALLOW_OFFICE_JS=0`).
- Tên tool PPT tránh trùng tên tool Word: `insert_slide_table`, `edit_slide_table`, `set_slide_hyperlink`. `distribute_shapes` gộp vào `align_shapes` (tham số `distribute`).
- `insert_toc`/`update_toc` dùng `document.tablesOfContents` (WordApiDesktop 1.4); `set_page_layout` cần WordApiDesktop 1.3 (desktop only).
- `read_paragraph_range` gộp vào `read_doc_section`; `read_reviewed_edits` không cần (kết quả trả thẳng từ `propose_doc_edits`); `read_attachment_pages` chưa làm (đã có `read_attached_file`).
- Word `ToolError` dùng `range_not_found` (không có `not_found` trong enum).
- Word/PPT: không auto-restore; chỉ backup.

### Chưa làm (phase 2 theo plan)

PPT: speaker notes, insert image/icon, chart, master, transition, sections, copy image giữa slide. P0 spike OOXML (notes/image/chart/master, đo `exportAsBase64`) **chưa chạy** vì cần Office thật; những giả định OOXML nằm ở báo cáo bàn giao.

### Cần smoke test live (checklist 4.4/4.6)

Mở Word, Excel, PowerPoint thật: ribbon, `hello` đúng host, 1 read + 1 write mỗi host, `edit_slide_xml`/`duplicate_slide` round-trip, `getFileAsync(Compressed)` trên Word/PPT web, track changes, comments, `propose_doc_edits`, TOC desktop.

### Cập nhật 09/11/2026 (rà soát trước khi lên production)

- Sửa test `tests/cli-update.test.ts` đỏ trên Linux: `bin/update.mjs` dùng `path.win32.resolve` cho đường dẫn Windows (trên Windows không đổi hành vi).
- Chữ "Excel" còn cứng trong giao diện pane (cài đặt, đăng nhập Copilot, bảng cập nhật, `execute_office_js`) nay theo host đang mở; `index.html` đánh dấu bằng `.host-app` và `applyHostBranding` điền tên app.
- Snapshot gửi model escape ký tự `<` (xem C10).
- `docs/install.md` thêm mục "Known limits for Word and PowerPoint"; `docs/manual-test.md` thêm checklist smoke test live.
- Mở lại pane (tắt rồi mở) luôn vào chat mới; chat cũ nằm trong lịch sử. Chỉ khi pane tự tải lại (build mới, cập nhật xong) mới mở lại chat đang dở: cờ `sessionStorage` trong `addin/taskpane/ui/chat-restore.ts`, đọc xong là xoá.
- Sửa các chỗ plan lệch code: đường dẫn `host-prompts`, `UndoJournal`, tên tool PPT, thẻ ngữ cảnh chưa tồn tại.

## 11. Việc còn lại trước khi gọi là production

Những mục dưới đây **không làm được trên máy không có Office** và là điều kiện chặn phát hành:

1. Smoke test live theo `docs/manual-test.md` (mục "Smoke test trực tiếp: Word và PowerPoint") trên Windows có đủ 3 app.
2. Xác nhận trên Office thật: ribbon hiện đúng một nút ở mỗi app; `edit_slide_xml` và `duplicate_slide` không làm file cần Repair; `getFileAsync(Compressed)` trên Word/PowerPoint (kể cả bản web).
3. Cài đè lên bản đang dùng (add-in `Id` giữ nguyên, version `1.4.0.0`) và kiểm tra không trùng nút ribbon.
4. Chọn số phiên bản npm cho đợt phát hành (hiện `package.json` vẫn là `0.1.21`) rồi chạy `npm run pack:release`; việc này làm khi chủ dự án quyết định.

## 12. Cross-file giữa 3 app, đọc file chưa mở và đổi tên gói

### Cross-file giữa Word, Excel và PowerPoint (đã làm, test giả lập)

- Mọi file đang mở với add-in thấy nhau trong danh sách (`ConnectedFile.host`); pane Word/PowerPoint cũng có nút Files.
- `read_connected_file` chạy **một tool đọc của chính app đó** (`CONNECTED_READ_TOOLS_BY_HOST` trong `protocol/src/tools-advanced.ts`): Excel như cũ; Word `get_doc_outline`, `read_doc_section`, `search_doc_text`, `read_comments`, `read_tracked_changes`; PowerPoint `get_presentation_metadata`, `list_slide_shapes`, `read_slide_text`. Tool của app khác bị từ chối trước khi hỏi người dùng. Không có tool ghi xuyên file.
- Mỗi lần đọc file khác vẫn phải được người dùng duyệt (một lần cho mỗi tin nhắn và mỗi file); hộp thoại nói rõ loại file (workbook/document/presentation).
- `send_to_connected_file` chạy được giữa mọi cặp app; bên nhận thấy "From <app> · <tên file>".
- Link ô chỉ có cho workbook; kết quả đọc từ Word/PowerPoint không kèm link.
- Test: `tests/cross-host-files.test.ts`, `tests/send-to-connected-file.test.ts`, `tests/connected-files*.test.ts`, `tests/pane-boot.test.ts`.

### Đọc file `.docx` / `.pptx` chưa mở (đã làm)

- Tool mới `read_document_file` (host-agnostic, bridge tự trả lời, không cần pane): `bridge/src/document-reader.ts`.
- Word: đoạn và bảng theo thứ tự đọc, style/cấp heading, dàn ý heading, comment; chữ tracked-change lấy bản đã chấp nhận. PowerPoint: tiêu đề, chữ, bảng, alt text ảnh và **speaker notes** (Office.js không đọc được notes). Phân trang bằng `from`/`nextFrom`, giới hạn 30.000 ký tự; mỗi phần XML giải nén tối đa 12 MB.
- Quyền đọc giống `read_attached_file`: thư mục tham chiếu, tệp đính kèm, hoặc hỏi người dùng cho đúng file đó.
- `read_workbook_file` cũng mở cho Word/PowerPoint để lấy số liệu từ file Excel trên đĩa.
- Test: `tests/document-reader.test.ts`, `tests/read-document-file.test.ts`; fixture `tests/support/office-files.ts`.

### Đổi tên gói npm `excel-partner` → `office-partner` (đã làm trong code, chưa publish)

- `package.json`: `name` `office-partner`, bin `office-partner` (+ `agent-gateway`). File CLI đổi thành `bin/office-partner.mjs`. Mọi chỉ dẫn lệnh trong thông báo, README và tài liệu dùng `office-partner`.
- Bộ cập nhật không còn ghi cứng tên: `bridge/src/updater.ts` và `bin/update.mjs` đọc `name` trong `package.json` của bản đang cài (URL registry, `npm install -g <tên>@latest`, thư mục global cần khớp). Biến môi trường `OFFICE_PARTNER_UPDATE_SPEC` (cũ: `EXCEL_PARTNER_UPDATE_SPEC` vẫn nhận).
- Giữ nguyên để không phá cài đặt cũ: thư mục dữ liệu `%LOCALAPPDATA%\Excel Partner`, `excel-partner.config.json`, `~/.excel-partner`, khoá khởi động `Excel Partner Bridge`, `Id` manifest.
- **Gói cầu nối** `legacy/excel-partner/` (publish dưới tên `excel-partner`, version `0.2.0`): chỉ có bộ chuyển lệnh. Người dùng cũ bấm Update (hoặc chạy `excel-partner`) → npm cài gói cầu nối → nó chạy `npm install -g office-partner` rồi chuyển lệnh sang `office-partner`. Không xoá gì (kể cả chính nó) nên không tranh chấp file trên Windows; không khai báo bin `agent-gateway` nên không đụng bin của `office-partner`.
- Đã thử bằng npm thật trên WSL (thư mục global tạm): cài `excel-partner@0.1.21` thật từ npm → cài đè gói cầu nối → chạy `excel-partner help` → tự cài `office-partner`, hiện đúng trợ giúp, lần sau không cài lại. Phát hiện và sửa: `npm root -g` che đường dẫn giống mã bí mật, nên cầu nối còn tìm ở thư mục của chính nó.
- Test: `tests/legacy-shim.test.ts`, `tests/updater.test.ts`, `tests/cli-update.test.ts`.

### Thứ tự phát hành (chưa làm, cần chủ dự án quyết định)

1. Đặt số phiên bản cho `office-partner` (hiện `package.json` vẫn `0.1.21`), `npm run build`, `npm publish`.
2. Kiểm tra `npm view office-partner version`, rồi `cd legacy/excel-partner && npm publish` (version `0.2.0`, phải cao hơn `0.1.21`).
3. Trên một máy Windows đang cài `excel-partner@0.1.21`: bấm **Update now**, rồi `office-partner status`. Mong đợi: bridge chạy, pane mở lại bình thường, `where office-partner` có kết quả.
4. Tuỳ chọn: `npm deprecate "excel-partner@<0.2.0" "Renamed to office-partner"`.
