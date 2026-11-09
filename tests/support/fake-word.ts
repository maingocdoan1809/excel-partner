/**
 * A small in-memory stand-in for the parts of Word's Office.js the Word executor uses. It is NOT Word: it checks that
 * the executor loads before it reads (like real proxies do), gates API sets through `isSetSupported`, enforces the
 * 255-character search limit and keeps live ranges across edits, but it applies every change immediately (no
 * batching). Install with `installWordFake()`; the returned FakeWordDoc is the document to arrange and inspect.
 */

export class WordError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
  }
}

/* ----- load-before-read proxies --------------------------------------------------------------------------------- */

const RAW = Symbol("raw");
interface Handle { commit(): void; markLoaded(names: string[]): void }
const handles = new WeakMap<object, Handle>();
const proxies = new WeakMap<object, object>();
const pendingLoads = new Set<Handle>();

function parseNames(names: unknown): string[] {
  const list = Array.isArray(names) ? names : typeof names === "string" ? names.split(",") : [];
  return list.map((n) => String(n).trim()).filter(Boolean);
}

const memos = new WeakMap<object, Map<string, unknown>>();
/** Navigation properties of a real Office.js object return the same proxy every time; so do these. */
function memo<T>(owner: object, key: string, make: () => T): T {
  let map = memos.get(owner);
  if (!map) { map = new Map(); memos.set(owner, map); }
  if (!map.has(key)) map.set(key, make());
  return map.get(key) as T;
}

function raw<T>(value: T): T {
  return ((value as unknown as Record<symbol, T>)?.[RAW] ?? value) as T;
}

function wrap<T extends object>(target: T, label: string, loadable: readonly string[], onCommit?: () => void): T {
  const existing = proxies.get(target);
  if (existing) return existing as T;
  const loaded = new Set<string>();
  const pending = new Set<string>();
  let itemNames: string[] = [];
  const handle: Handle = {
    commit() {
      for (const n of pending) loaded.add(n);
      pending.clear();
      onCommit?.();
      if (loaded.has("items") && itemNames.length > 0) {
        for (const item of (target as unknown as { _items: object[] })._items) handles.get(raw(item))?.markLoaded(itemNames);
      }
    },
    markLoaded(names) {
      for (const n of names) loaded.add(n);
    },
  };
  handles.set(target, handle);
  const proxy: T = new Proxy(target, {
    get(t, key) {
      if (key === RAW) return t;
      if (key === "load") {
        return (names?: unknown) => {
          const requested = names === undefined ? [...loadable] : parseNames(names);
          for (const name of requested) {
            if (name.startsWith("items/")) {
              pending.add("items");
              itemNames = [...new Set([...itemNames, name.slice(6)])];
            } else pending.add(name);
          }
          pendingLoads.add(handle);
          return proxy;
        };
      }
      if (typeof key === "string" && loadable.includes(key) && !loaded.has(key)) {
        throw new WordError("PropertyNotLoaded", `The property '${key}' is not available on this ${label}. Before reading the property's value, call the load method on the containing object and call "context.sync()".`);
      }
      const value = Reflect.get(t, key, t);
      return typeof value === "function" ? value.bind(t) : value;
    },
    set(t, key, value) {
      return Reflect.set(t, key, value, t);
    },
  });
  proxies.set(target, proxy);
  return proxy;
}

function commitLoads(): void {
  const all = [...pendingLoads];
  pendingLoads.clear();
  for (const handle of all) handle.commit();
}

/* ----- document model -------------------------------------------------------------------------------------------- */

export interface ParaData {
  text: string;
  style: string;
  styleBuiltIn: string;
  tableNestingLevel: number;
  pictures: number;
  listId: number | null;
  level: number;
  alignment: string;
  spaceAfter: number;
  spaceBefore: number;
  lineSpacing: number;
  font: Record<string, unknown>;
}

export interface CommentData {
  id: string;
  content: string;
  authorName: string;
  creationDate: Date;
  resolved: boolean;
  replies: Array<{ id: string; content: string; authorName: string; creationDate: Date }>;
  anchor: FRange;
}

export interface TrackedData { type: string; author: string; date: Date; text: string }
export interface FieldData { type: string; code: string; result: string; updates: number }
export interface ControlData { id: number; tag: string; title: string; text: string; type: string }
export interface TableData { rows: string[][]; style: string; headerRowCount: number; anchor: ParaData | null; merges: number[][] }

type Pos = [number, number];

export class FakeWordDoc {
  paragraphs: ParaData[] = [];
  comments: CommentData[] = [];
  tracked: TrackedData[] = [];
  trackBaseline: string | null = null;
  changeTrackingMode = "Off";
  tables: TableData[] = [];
  fields: FieldData[] = [];
  controls: Array<ControlData & { para: ParaData | null }> = [];
  bookmarks = new Map<string, string>();
  hyperlinks: Array<{ text: string; url: string }> = [];
  notes: Array<{ kind: string; text: string; after: string }> = [];
  tocs: Array<{ options: unknown; updatedPageNumbers: number }> = [];
  htmlInserts: Array<{ html: string; location: string }> = [];
  lists = new Map<number, { bullets: Record<number, string>; numbering: Record<number, string> }>();
  sections = [{
    header: { Primary: "", FirstPage: "", EvenPages: "" } as Record<string, string>,
    footer: { Primary: "", FirstPage: "", EvenPages: "" } as Record<string, string>,
    page: { orientation: "Portrait", paperSize: "Letter", topMargin: 72, bottomMargin: 72, leftMargin: 72, rightMargin: 72, pageWidth: 612, pageHeight: 792 } as Record<string, unknown>,
  }];
  url = "";
  sets: Record<string, string> = { WordApi: "1.6", WordApiDesktop: "1.4" };
  ranges = new Set<FRange>();
  selection: FRange;
  syncs = 0;
  /** Index of each whole paragraph deleted, in call order. */
  deleteLog: number[] = [];
  private nextId = 1;
  private nextListId = 1;

  constructor(paragraphs: string[] = []) {
    paragraphs.forEach((text) => this.paragraphs.push(this.newPara(text)));
    if (this.paragraphs.length === 0) this.paragraphs.push(this.newPara(""));
    this.selection = new FRange(this, this.paragraphs[0]!, 0, this.paragraphs[0]!, 0);
  }

  uid(prefix: string): string {
    return `${prefix}${this.nextId++}`;
  }
  newListId(): number {
    this.lists.set(this.nextListId, { bullets: {}, numbering: {} });
    return this.nextListId++;
  }

  newPara(text: string, style = "Normal", extra: Partial<ParaData> = {}): ParaData {
    const builtIn = /^Heading (\d)$/.exec(style) ? style.replace(" ", "") : style === "Title" ? "Title" : style === "Normal" ? "Normal" : "Other";
    return { text, style, styleBuiltIn: builtIn, tableNestingLevel: 0, pictures: 0, listId: null, level: 0, alignment: "Left", spaceAfter: 0, spaceBefore: 0, lineSpacing: 12, font: {}, ...extra };
  }

  /** Replace the paragraphs; each entry is plain text, `[text, style]` or a ready ParaData. */
  setParagraphs(entries: Array<string | [string, string] | Partial<ParaData> & { text: string }>): void {
    this.paragraphs = entries.map((entry) => typeof entry === "string" ? this.newPara(entry) : Array.isArray(entry) ? this.newPara(entry[0], entry[1]) : this.newPara(entry.text, entry.style ?? "Normal", entry));
    this.selection = new FRange(this, this.paragraphs[0]!, 0, this.paragraphs[0]!, 0);
  }

  texts(): string[] {
    return this.paragraphs.map((p) => p.text);
  }

  /** Put the selection on `text` (first occurrence) or at a cursor position in paragraph `index`. */
  select(index: number, start = 0, end = start): void {
    this.ranges.delete(this.selection);
    this.selection = new FRange(this, this.paragraphs[index]!, start, this.paragraphs[index]!, end);
  }

  supported(set: string, version: string): boolean {
    const have = this.sets[set];
    if (!have) return false;
    const a = have.split(".").map(Number);
    const b = version.split(".").map(Number);
    return a[0]! > b[0]! || (a[0] === b[0] && (a[1] ?? 0) >= (b[1] ?? 0));
  }

  requires(set: string, version: string, what: string): void {
    if (!this.supported(set, version)) throw new WordError("ApiNotFound", `${what} is not available (needs ${set} ${version}).`);
  }

  idx(p: ParaData): number {
    return this.paragraphs.indexOf(p);
  }
  pos(p: ParaData, off: number): Pos {
    return [this.idx(p), off];
  }

  fullText(): string {
    return this.paragraphs.map((p) => p.text).join("\r");
  }

  orderedTables(): TableData[] {
    const key = (t: TableData) => (t.anchor ? this.idx(t.anchor) : Number.MAX_SAFE_INTEGER);
    return [...this.tables].sort((a, b) => key(a) - key(b));
  }
}

function cmp(a: Pos, b: Pos): number {
  return a[0] - b[0] || a[1] - b[1];
}

/* ----- ranges ------------------------------------------------------------------------------------------------------ */

export class FRange {
  whole = false;
  constructor(readonly doc: FakeWordDoc, public p1: ParaData, public s: number, public p2: ParaData, public e: number) {
    doc.ranges.add(this);
  }

  get start(): Pos { return this.doc.pos(this.p1, this.s); }
  get end(): Pos { return this.doc.pos(this.p2, this.e); }

  get text(): string {
    if (this.whole) return this.paras().map((p) => p.text).join("\r");
    if (this.p1 === this.p2) return this.p1.text.slice(this.s, this.e);
    const ps = this.paras();
    return [this.p1.text.slice(this.s), ...ps.slice(1, -1).map((p) => p.text), this.p2.text.slice(0, this.e)].join("\r");
  }

  paras(): ParaData[] {
    const [a] = this.start;
    const [b] = this.end;
    return this.doc.paragraphs.slice(a, b + 1);
  }

  get style(): string { return this.p1.style; }
  set style(value: string) { for (const p of this.paras()) { p.style = value; p.styleBuiltIn = "Other"; } }
  set styleBuiltIn(value: string) { for (const p of this.paras()) { this.doc.requires("WordApi", "1.3", "styleBuiltIn"); p.styleBuiltIn = value; p.style = value.replace(/^Heading(\d)$/, "Heading $1"); } }
  get font(): object { return memo(this, "font", () => wrap(fontOf(this), "Font", ["highlightColor"])); }
  get paragraphs(): object { return memo(this, "paragraphs", () => collection(() => this.paras().map(paraWrap(this.doc)), "ParagraphCollection")); }
  set hyperlink(url: string) {
    this.doc.requires("WordApi", "1.3", "Range.hyperlink");
    this.doc.hyperlinks.push({ text: this.text, url });
  }

  getRange(location = "Whole"): FRange {
    const d = this.doc;
    switch (location) {
      case "Start": return rw(new FRange(d, this.p1, this.s, this.p1, this.s));
      case "End":
      case "After": return rw(new FRange(d, this.p2, this.e, this.p2, this.e));
      default: return rw(this);
    }
  }

  expandTo(other: FRange): FRange {
    const a = cmp(this.start, other.start) <= 0 ? this : other;
    const b = cmp(this.end, other.end) >= 0 ? this : other;
    return rw(new FRange(this.doc, a.p1, a.s, b.p2, b.e));
  }

  search(text: string, options: { matchCase?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean } = {}): object {
    return collection(() => searchHits(this.doc, text, options, this).map((r) => wrap(r, "Range", ["text", "style"])), "RangeCollection");
  }

  insertText(text: string, location: string): FRange {
    if (location === "Replace") { replaceRange(this.doc, this, text); return rw(this); }
    const point = location === "Start" || location === "Before" ? this.getRange("Start") : this.getRange("End");
    replaceRange(this.doc, point, text);
    return point;
  }
  delete(): void {
    if (this.whole && this.doc.paragraphs.length > 1) {
      const [a] = this.start;
      const [b] = this.end;
      this.doc.deleteLog.push(a);
      for (const q of this.doc.comments) if (q.anchor.start[0] >= a && q.anchor.end[0] <= b) this.doc.comments.splice(this.doc.comments.indexOf(q), 1);
      this.doc.paragraphs.splice(a, b - a + 1);
      if (this.doc.paragraphs.length === 0) this.doc.paragraphs.push(this.doc.newPara(""));
      return;
    }
    replaceRange(this.doc, this, "");
  }
  insertHtml(html: string, location: string): FRange {
    this.doc.htmlInserts.push({ html, location });
    const plain = html.replace(/<[^>]+>/g, "|").split("|").filter(Boolean).join(" ");
    return this.insertText(plain, location === "Replace" ? "Replace" : location);
  }
  insertParagraph(text: string, location: string): object {
    return paraInsert(this.doc, this.p2, text, location);
  }
  insertComment(text: string): object {
    this.doc.requires("WordApi", "1.4", "Range.insertComment");
    const data: CommentData = { id: this.doc.uid("c"), content: text, authorName: "Tester", creationDate: new Date("2026-01-02T03:04:05Z"), resolved: false, replies: [], anchor: this.cloneRange() };
    this.doc.comments.push(data);
    return commentWrap(this.doc)(data);
  }
  cloneRange(): FRange { return new FRange(this.doc, this.p1, this.s, this.p2, this.e); }
  getComments(): object {
    this.doc.requires("WordApi", "1.4", "Range.getComments");
    const mine = () => this.doc.comments.filter((c) => cmp(c.anchor.start, this.end) <= 0 && cmp(c.anchor.end, this.start) >= 0).map(commentWrap(this.doc));
    return collection(mine, "CommentCollection");
  }
  getTrackedChanges(): object {
    this.doc.requires("WordApi", "1.6", "Range.getTrackedChanges");
    return trackedCollection(this.doc);
  }
  insertFootnote(text: string): object { return this.note("footnote", text); }
  insertEndnote(text: string): object { return this.note("endnote", text); }
  private note(kind: string, text: string): object {
    this.doc.requires("WordApi", "1.5", `Range.insert${kind === "footnote" ? "Footnote" : "Endnote"}`);
    this.doc.notes.push({ kind, text, after: this.text });
    return {};
  }
  insertBookmark(name: string): void {
    this.doc.requires("WordApi", "1.4", "Range.insertBookmark");
    this.doc.bookmarks.set(name, this.text);
  }
  getBookmarks(): { value: string[] } {
    this.doc.requires("WordApi", "1.4", "Range.getBookmarks");
    return { value: [...this.doc.bookmarks.keys()] };
  }
  insertField(location: string, type: string, text?: string): object {
    this.doc.requires("WordApi", "1.5", "Range.insertField");
    if (!["Replace", "Start", "End", "Before", "After"].includes(location)) throw new WordError("InvalidArgument", "bad location");
    const data: FieldData = { type, code: `${type}${text ? ` ${text}` : ""}`, result: `‹${type}›`, updates: 0 };
    this.doc.fields.push(data);
    return fieldWrap(this.doc)(data);
  }
  insertContentControl(): object {
    const data = { id: Number(this.doc.uid("")), tag: "", title: "", text: this.text, type: "RichText", para: this.p1 };
    this.doc.controls.push(data);
    return controlWrap(this.doc)(data);
  }
  get contentControls(): object { return collection(() => this.doc.controls.map(controlWrap(this.doc)), "ContentControlCollection"); }
  insertTable(rows: number, columns: number, location: string, values?: string[][]): object {
    return addTable(this.doc, rows, columns, values, this.p2);
  }
  select(): void { this.doc.selection = this; }
}

const rw = (r: FRange): FRange => wrap(r, "Range", ["text", "style"]);

function fontOf(owner: { paras(): ParaData[] }): object {
  const paras = () => owner.paras();
  return {
    get highlightColor() { return (paras()[0]!.font.highlightColor as string | null | undefined) ?? null; },
    set highlightColor(value: string | null) { for (const p of paras()) p.font.highlightColor = value; },
    set name(v: string) { for (const p of paras()) p.font.name = v; },
    set size(v: number) { for (const p of paras()) p.font.size = v; },
    set bold(v: boolean) { for (const p of paras()) p.font.bold = v; },
    set italic(v: boolean) { for (const p of paras()) p.font.italic = v; },
    set underline(v: string) { for (const p of paras()) p.font.underline = v; },
    set color(v: string) { for (const p of paras()) p.font.color = v; },
  };
}

/* ----- search + replace ------------------------------------------------------------------------------------------- */

function searchHits(doc: FakeWordDoc, text: string, options: { matchCase?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean }, scope?: FRange): FRange[] {
  if (text.length > 255) throw new WordError("SearchStringInvalidOrTooLong", "The search text is over 255 characters.");
  if (options.matchWildcards) throw new WordError("NotImplemented", "wildcards are not supported by the fake");
  let needle = "";
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "^") {
      const next = text[i + 1];
      if (next === "^") { needle += "^"; i++; } else if (next === "p") { needle += "\r"; i++; } else if (next === "t") { needle += "\t"; i++; } else throw new WordError("InvalidArgument", `Unknown search code ^${next ?? ""}`);
    } else needle += text[i];
  }
  if (needle === "") return [];
  const hay = doc.fullText();
  const h = options.matchCase ? hay : hay.toLowerCase();
  const n = options.matchCase ? needle : needle.toLowerCase();
  const starts: number[] = [];
  let at = h.indexOf(n);
  while (at !== -1) {
    const before = hay[at - 1];
    const after = hay[at + n.length];
    const word = (c: string | undefined) => c !== undefined && /[\p{L}\p{N}_]/u.test(c);
    if (!options.matchWholeWord || (!word(before) && !word(after))) starts.push(at);
    at = h.indexOf(n, at + n.length);
  }
  const locate = (offset: number, atEnd: boolean): [ParaData, number] => {
    let base = 0;
    for (const p of doc.paragraphs) {
      const limit = base + p.text.length;
      if (offset < limit || (atEnd && offset === limit)) return [p, offset - base];
      base = limit + 1;
    }
    const last = doc.paragraphs[doc.paragraphs.length - 1]!;
    return [last, last.text.length];
  };
  const hits = starts.map((s) => {
    const [p1, o1] = locate(s, false);
    const [p2, o2] = locate(s + n.length, true);
    return new FRange(doc, p1, o1, p2, o2);
  });
  return scope ? hits.filter((r) => cmp(r.start, scope.start) >= 0 && cmp(r.end, scope.end) <= 0) : hits;
}

function replaceRange(doc: FakeWordDoc, range: FRange, newText: string): void {
  const r = raw(range);
  const [i1] = r.start;
  const [i2] = r.end;
  const removed = r.text;
  const prefix = r.p1.text.slice(0, r.s);
  const suffix = r.p2.text.slice(r.e);
  const oldP2 = r.p2;
  const oldE = r.e;
  const oldStart = r.start;
  const oldEnd = r.end;
  const before = new Map<FRange, { a: Pos; b: Pos }>();
  for (const q of doc.ranges) before.set(q, { a: q.start, b: q.end });
  const stamp = new Date("2026-01-02T03:04:05Z");
  if (doc.changeTrackingMode !== "Off" && (removed !== "" || newText !== "")) {
    if (doc.trackBaseline === null) doc.trackBaseline = doc.fullText();
    if (removed !== "") doc.tracked.push({ type: "Deleted", author: "Tester", date: stamp, text: removed });
    if (newText !== "") doc.tracked.push({ type: "Added", author: "Tester", date: stamp, text: newText });
  }
  // Comments anchored inside the replaced text go away with it.
  if (cmp(oldStart, oldEnd) !== 0) {
    doc.comments = doc.comments.filter((c) => {
      const span = before.get(c.anchor)!;
      return !(c.anchor !== r && cmp(span.a, oldStart) >= 0 && cmp(span.b, oldEnd) <= 0);
    });
  }
  const parts = newText.split(/\r\n|\n|\r/);
  let lastPara: ParaData;
  let lastLen: number;
  if (parts.length === 1) {
    r.p1.text = prefix + parts[0] + suffix;
    doc.paragraphs.splice(i1 + 1, i2 - i1);
    lastPara = r.p1;
    lastLen = prefix.length + parts[0]!.length;
  } else {
    r.p1.text = prefix + parts[0];
    const created = parts.slice(1).map((t, k) => doc.newPara(k === parts.length - 2 ? t + suffix : t, r.p1.style, { styleBuiltIn: r.p1.styleBuiltIn }));
    doc.paragraphs.splice(i1 + 1, i2 - i1, ...created);
    lastPara = created[created.length - 1]!;
    lastLen = parts[parts.length - 1]!.length;
  }
  const firstPara = r.p1;
  const firstOff = oldStart[1];
  for (const q of doc.ranges) {
    if (q === r) continue;
    const was = before.get(q)!;
    const shift = (p: ParaData, off: number): [ParaData, number] => (p === oldP2 ? [lastPara, lastLen + (off - oldE)] : [p, off]);
    if (cmp(was.a, oldStart) > 0) {
      if (cmp(was.a, oldEnd) >= 0) [q.p1, q.s] = shift(q.p1, q.s);
      else [q.p1, q.s] = [firstPara, firstOff];
    }
    if (cmp(was.b, oldStart) > 0) {
      if (cmp(was.b, oldEnd) > 0) [q.p2, q.e] = shift(q.p2, q.e);
      else [q.p2, q.e] = [lastPara, lastLen];
    }
  }
  r.p2 = lastPara;
  r.e = lastLen;
  r.whole = false;
}

/* ----- collections + factories ------------------------------------------------------------------------------------- */

function collection(getItems: () => object[], label: string, extra: Record<string, unknown> = {}): any {
  const target: Record<string, unknown> & { _items: object[] } = {
    _items: [],
    get items() { return this._items; },
    getFirst() { const first = getItems()[0]; if (!first) throw new WordError("ItemNotFound", `${label} is empty`); return first; },
    getLast() { const last = getItems().at(-1); if (!last) throw new WordError("ItemNotFound", `${label} is empty`); return last; },
    ...extra,
  };
  return wrap(target, label, ["items"], () => { target._items = getItems(); });
}

const paraCache = new WeakMap<ParaData, object>();
function paraWrap(doc: FakeWordDoc): (data: ParaData) => object {
  return (data) => {
    const cached = paraCache.get(data);
    if (cached) return cached;
    const p = new FPara(doc, data);
    const proxy = wrap(p, "Paragraph", ["text", "style", "styleBuiltIn", "tableNestingLevel", "isListItem", "alignment", "spaceAfter", "spaceBefore", "lineSpacing"]);
    paraCache.set(data, proxy);
    return proxy;
  };
}

class FPara {
  constructor(readonly doc: FakeWordDoc, readonly data: ParaData) {}
  get text() { return this.data.text; }
  get style() { return this.data.style; }
  set style(v: string) { this.data.style = v; this.data.styleBuiltIn = "Other"; }
  get styleBuiltIn() { return this.data.styleBuiltIn; }
  set styleBuiltIn(v: string) { this.doc.requires("WordApi", "1.3", "Paragraph.styleBuiltIn"); this.data.styleBuiltIn = v; this.data.style = v.replace(/^Heading(\d)$/, "Heading $1"); }
  get tableNestingLevel() { this.doc.requires("WordApi", "1.3", "tableNestingLevel"); return this.data.tableNestingLevel; }
  get isListItem() { this.doc.requires("WordApi", "1.3", "isListItem"); return this.data.listId !== null; }
  get alignment() { return this.data.alignment; }
  set alignment(v: string) { this.data.alignment = v; }
  get spaceAfter() { return this.data.spaceAfter; }
  set spaceAfter(v: number) { this.data.spaceAfter = v; }
  get spaceBefore() { return this.data.spaceBefore; }
  set spaceBefore(v: number) { this.data.spaceBefore = v; }
  get lineSpacing() { return this.data.lineSpacing; }
  set lineSpacing(v: number) { this.data.lineSpacing = v; }
  paras(): ParaData[] { return [this.data]; }
  get font() { return memo(this, "font", () => wrap(fontOf(this), "Font", ["highlightColor"])); }
  get inlinePictures() { return memo(this, "pictures", () => collection(() => Array.from({ length: this.data.pictures }, () => ({})), "InlinePictureCollection")); }
  get contentControls() { return memo(this, "controls", () => collection(() => this.doc.controls.filter((c) => c.para === this.data).map(controlWrap(this.doc)), "ContentControlCollection")); }
  get listItem() {
    const data = this.data;
    return { get level() { return data.level; }, set level(v: number) { data.level = v; } };
  }
  getRange(location = "Whole"): FRange {
    const d = this.doc;
    const r = new FRange(d, this.data, 0, this.data, this.data.text.length);
    if (location === "Whole") r.whole = true;
    if (location === "Start") r.e = 0;
    if (location === "End" || location === "After") r.s = r.e;
    return rw(r);
  }
  delete(): void { this.getRange("Whole").delete(); }
  insertParagraph(text: string, location: string): object { return paraInsert(this.doc, this.data, text, location); }
  insertHtml(html: string, location: string): FRange { return this.getRange("Whole").insertHtml(html, location === "Replace" ? "Replace" : location); }
  insertTable(rows: number, columns: number, location: string, values?: string[][]): object { return addTable(this.doc, rows, columns, values, this.data); }
  startNewList(): object {
    this.doc.requires("WordApi", "1.3", "startNewList");
    const id = this.doc.newListId();
    this.data.listId = id;
    return listWrap(this.doc, id);
  }
  attachToList(id: number, level: number): object {
    this.doc.requires("WordApi", "1.3", "attachToList");
    if (!this.doc.lists.has(id)) throw new WordError("ItemNotFound", `no list ${id}`);
    this.data.listId = id;
    this.data.level = level;
    return listWrap(this.doc, id);
  }
}

function paraInsert(doc: FakeWordDoc, anchor: ParaData, text: string, location: string): object {
  const created = doc.newPara(text, anchor.style, { styleBuiltIn: anchor.styleBuiltIn });
  const at = doc.paragraphs.indexOf(anchor);
  if (location === "After") doc.paragraphs.splice(at + 1, 0, created);
  else if (location === "Before") doc.paragraphs.splice(at, 0, created);
  else throw new WordError("InvalidArgument", `bad location ${location}`);
  return paraWrap(doc)(created);
}

function listWrap(doc: FakeWordDoc, id: number): object {
  const data = doc.lists.get(id)!;
  return wrap({
    get id() { return id; },
    setLevelBullet(level: number, bullet: string) { doc.requires("WordApi", "1.3", "setLevelBullet"); data.bullets[level] = bullet; },
    setLevelNumbering(level: number, numbering: string) { doc.requires("WordApi", "1.3", "setLevelNumbering"); data.numbering[level] = numbering; },
  }, "List", ["id"]);
}

function addTable(doc: FakeWordDoc, rows: number, columns: number, values: string[][] | undefined, anchor: ParaData | null): object {
  doc.requires("WordApi", "1.3", "insertTable");
  const grid = Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => values?.[r]?.[c] ?? ""));
  const data: TableData = { rows: grid, style: "Table Grid", headerRowCount: 0, anchor, merges: [] };
  doc.tables.push(data);
  return tableWrap(doc)(data);
}

const tableCache = new WeakMap<TableData, object>();
function tableWrap(doc: FakeWordDoc): (data: TableData) => object {
  return (data) => {
    const hit = tableCache.get(data);
    if (hit) return hit;
    const t = {
      get rowCount() { return data.rows.length; },
      get values() { return data.rows.map((r) => [...r]); },
      get headerRowCount() { return data.headerRowCount; },
      set headerRowCount(v: number) { data.headerRowCount = v; },
      get style() { return data.style; },
      set style(v: string) { if (v.startsWith("Bad")) throw new WordError("ItemNotFound", `style ${v} not found`); data.style = v; },
      get rows() {
        return memo(this, "rows", () => collection(() => data.rows.map((_, i) => wrap({
          get cellCount() { return data.rows[i]?.length ?? 0; },
          insertRows(loc: string, count: number) { doc.requires("WordApi", "1.3", "insertRows"); data.rows.splice(loc === "Before" ? i : i + 1, 0, ...Array.from({ length: count }, () => data.rows[0]!.map(() => ""))); return {}; },
        }, "TableRow", ["cellCount"])), "TableRowCollection"));
      },
      addRows(loc: string, count: number) { data.rows.splice(loc === "Start" ? 0 : data.rows.length, 0, ...Array.from({ length: count }, () => data.rows[0]!.map(() => ""))); return {}; },
      deleteRows(index: number, count = 1) { data.rows.splice(index, count); },
      addColumns(loc: string, count: number) { for (const r of data.rows) r.splice(loc === "Start" ? 0 : r.length, 0, ...Array.from({ length: count }, () => "")); },
      deleteColumns(index: number, count = 1) { for (const r of data.rows) r.splice(index, count); },
      mergeCells(a: number, b: number, c: number, d: number) { doc.requires("WordApi", "1.4", "mergeCells"); data.merges.push([a, b, c, d]); return {}; },
      delete() { doc.tables.splice(doc.tables.indexOf(data), 1); },
      getCell(row: number, col: number) {
        if (row >= data.rows.length || col >= (data.rows[0]?.length ?? 0)) throw new WordError("InvalidArgument", "cell out of range");
        return wrap({
          get value() { return data.rows[row]![col]!; },
          set value(v: string) { data.rows[row]![col] = v; },
          insertColumns(loc: string, count: number) { for (const r of data.rows) r.splice(loc === "Before" ? col : col + 1, 0, ...Array.from({ length: count }, () => "")); },
        }, "TableCell", ["value"]);
      },
    };
    const proxy = wrap(t, "Table", ["rowCount", "values", "headerRowCount", "style"]);
    tableCache.set(data, proxy);
    return proxy;
  };
}

const commentCache = new WeakMap<CommentData, object>();
function commentWrap(doc: FakeWordDoc): (data: CommentData) => object {
  return (data) => {
    const hit = commentCache.get(data);
    if (hit) return hit;
    const c = {
      get id() { return data.id; },
      get content() { return data.content; },
      get authorName() { return data.authorName; },
      get creationDate() { return data.creationDate; },
      get resolved() { return data.resolved; },
      set resolved(v: boolean) { data.resolved = v; },
      get replies() {
        return memo(this, "replies", () => collection(() => data.replies.map((r) => wrap({ get id() { return r.id; }, get content() { return r.content; }, get authorName() { return r.authorName; }, get creationDate() { return r.creationDate; } }, "CommentReply", ["id", "content", "authorName", "creationDate"])), "CommentReplyCollection"));
      },
      getRange() { return wrap(data.anchor, "Range", ["text", "style"]); },
      reply(text: string) {
        const reply = { id: doc.uid("r"), content: text, authorName: "Tester", creationDate: new Date("2026-01-02T03:04:06Z") };
        data.replies.push(reply);
        return wrap({ get id() { return reply.id; } }, "CommentReply", ["id"]);
      },
    };
    const proxy = wrap(c, "Comment", ["id", "content", "authorName", "creationDate", "resolved"]);
    commentCache.set(data, proxy);
    return proxy;
  };
}

function trackedCollection(doc: FakeWordDoc): object {
  const wrapOne = (data: TrackedData) => wrap({
    get author() { return data.author; },
    get date() { return data.date; },
    get text() { return data.text; },
    get type() { return data.type; },
    accept() { doc.tracked.splice(doc.tracked.indexOf(data), 1); },
    reject() { doc.tracked.splice(doc.tracked.indexOf(data), 1); },
  }, "TrackedChange", ["author", "date", "text", "type"]);
  const cache = new Map<TrackedData, object>();
  return collection(() => doc.tracked.map((d) => { if (!cache.has(d)) cache.set(d, wrapOne(d)); return cache.get(d)!; }), "TrackedChangeCollection", {
    acceptAll() { doc.tracked.length = 0; },
    rejectAll() { doc.tracked.length = 0; },
  });
}

const fieldCache = new WeakMap<FieldData, object>();
function fieldWrap(doc: FakeWordDoc): (data: FieldData) => object {
  return (data) => {
    const hit = fieldCache.get(data);
    if (hit) return hit;
    const f = {
      get code() { doc.requires("WordApi", "1.4", "Field.code"); return data.code; },
      get type() { return data.type; },
      get result() { return memo(this, "result", () => wrap({ get text() { return data.result; } }, "Range", ["text"])); },
      updateResult() { doc.requires("WordApi", "1.5", "Field.updateResult"); if (data.type === "Bad") throw new WordError("GeneralException", "cannot update"); data.updates++; },
    };
    const proxy = wrap(f, "Field", ["code", "type"]);
    fieldCache.set(data, proxy);
    return proxy;
  };
}

const controlCache = new WeakMap<object, object>();
function controlWrap(doc: FakeWordDoc): (data: ControlData & { para: ParaData | null }) => object {
  return (data) => {
    const hit = controlCache.get(data);
    if (hit) return hit;
    const c = {
      get id() { return data.id; },
      get tag() { return data.tag; },
      set tag(v: string) { data.tag = v; },
      get title() { return data.title; },
      set title(v: string) { data.title = v; },
      get text() { return data.text; },
      get type() { return data.type; },
      insertText(text: string) { data.text = text; return {}; },
      delete(keepContent: boolean) { doc.controls.splice(doc.controls.indexOf(data), 1); void keepContent; },
    };
    const proxy = wrap(c, "ContentControl", ["id", "tag", "title", "text", "type"]);
    controlCache.set(data, proxy);
    return proxy;
  };
}

/* ----- body, document, context ------------------------------------------------------------------------------------ */

function headerBody(doc: FakeWordDoc, bag: Record<string, string>, type: string): object {
  return wrap({
    get text() { return bag[type]!; },
    insertText(text: string) { bag[type] = text; return {}; },
  }, "Body", ["text"]);
}

const bodies = new WeakMap<FakeWordDoc, object>();
function makeBody(doc: FakeWordDoc): object {
  const hit = bodies.get(doc);
  if (hit) return hit;
  const made = buildBody(doc);
  bodies.set(doc, made);
  return made;
}

function buildBody(doc: FakeWordDoc): object {
  const whole = () => rw(new FRange(doc, doc.paragraphs[0]!, 0, doc.paragraphs.at(-1)!, doc.paragraphs.at(-1)!.text.length));
  const body = {
    get paragraphs() { return memo(this, "paragraphs", () => collection(() => doc.paragraphs.map(paraWrap(doc)), "ParagraphCollection")); },
    get tables() { doc.requires("WordApi", "1.3", "Body.tables"); return memo(this, "tables", () => collection(() => doc.orderedTables().map(tableWrap(doc)), "TableCollection")); },
    get fields() { doc.requires("WordApi", "1.4", "Body.fields"); return memo(this, "fields", () => collection(() => doc.fields.map(fieldWrap(doc)), "FieldCollection")); },
    get contentControls() {
      return memo(this, "controls", () => collection(() => doc.controls.map(controlWrap(doc)), "ContentControlCollection", {
        getByTag: (tag: string) => collection(() => doc.controls.filter((c) => c.tag === tag).map(controlWrap(doc)), "ContentControlCollection"),
        getByTitle: (title: string) => collection(() => doc.controls.filter((c) => c.title === title).map(controlWrap(doc)), "ContentControlCollection"),
      }));
    },
    search(text: string, options?: object) { return whole().search(text, options); },
    getComments() { doc.requires("WordApi", "1.4", "Body.getComments"); return collection(() => doc.comments.map(commentWrap(doc)), "CommentCollection"); },
    getTrackedChanges() { doc.requires("WordApi", "1.6", "Body.getTrackedChanges"); return trackedCollection(doc); },
    getReviewedText(version = "Current") {
      doc.requires("WordApi", "1.4", "Body.getReviewedText");
      return { value: version === "Original" && doc.trackBaseline !== null ? doc.trackBaseline : doc.fullText() };
    },
    getRange(location = "Whole") {
      doc.requires("WordApi", "1.3", "Body.getRange");
      const w = whole();
      return location === "End" ? w.getRange("End") : location === "Start" ? w.getRange("Start") : w;
    },
    insertParagraph(text: string, location: string) {
      const created = doc.newPara(text);
      if (location === "End") doc.paragraphs.push(created);
      else if (location === "Start") doc.paragraphs.unshift(created);
      else throw new WordError("InvalidArgument", `bad location ${location}`);
      return paraWrap(doc)(created);
    },
    insertTable(rows: number, columns: number, location: string, values?: string[][]) { return addTable(doc, rows, columns, values, null); },
    insertHtml(html: string, location: string) { doc.htmlInserts.push({ html, location }); doc.paragraphs.push(doc.newPara(html.replace(/<[^>]+>/g, " ").trim())); return {}; },
  };
  return wrap(body, "Body", []);
}

function makeDocument(doc: FakeWordDoc): object {
  const sections = () => doc.sections.map((s) => wrap({
    getHeader: (type: string) => headerBody(doc, s.header, type),
    getFooter: (type: string) => headerBody(doc, s.footer, type),
    get pageSetup() {
      doc.requires("WordApiDesktop", "1.3", "Section.pageSetup");
      return wrap({
        get orientation() { return s.page.orientation as string; },
        set orientation(v: string) { if (v !== s.page.orientation) { const w = s.page.pageWidth; s.page.pageWidth = s.page.pageHeight; s.page.pageHeight = w; } s.page.orientation = v; },
        get paperSize() { return s.page.paperSize as string; },
        set paperSize(v: string) { s.page.paperSize = v; },
        get topMargin() { return s.page.topMargin as number; }, set topMargin(v: number) { s.page.topMargin = v; },
        get bottomMargin() { return s.page.bottomMargin as number; }, set bottomMargin(v: number) { s.page.bottomMargin = v; },
        get leftMargin() { return s.page.leftMargin as number; }, set leftMargin(v: number) { s.page.leftMargin = v; },
        get rightMargin() { return s.page.rightMargin as number; }, set rightMargin(v: number) { s.page.rightMargin = v; },
        get pageWidth() { return s.page.pageWidth as number; },
        get pageHeight() { return s.page.pageHeight as number; },
      }, "PageSetup", ["orientation", "paperSize", "topMargin", "bottomMargin", "leftMargin", "rightMargin", "pageWidth", "pageHeight"]);
    },
  }, "Section", []));
  const document = {
    get body() { return makeBody(doc); },
    get changeTrackingMode() { doc.requires("WordApi", "1.4", "changeTrackingMode"); return doc.changeTrackingMode; },
    set changeTrackingMode(v: string) { doc.requires("WordApi", "1.4", "changeTrackingMode"); doc.changeTrackingMode = v; },
    get sections() { return collection(sections, "SectionCollection"); },
    get tablesOfContents() {
      doc.requires("WordApiDesktop", "1.4", "Document.tablesOfContents");
      return collection(() => doc.tocs.map((t) => wrap({
        get range() {
          return wrap({
            get fields() { return collection(() => [fieldWrap(doc)({ type: "TOC", code: "TOC", result: "", updates: 0 })], "FieldCollection"); },
          }, "Range", []);
        },
        updatePageNumbers() { t.updatedPageNumbers++; },
      }, "TableOfContents", [])), "TableOfContentsCollection", {
        add(range: FRange, options: unknown) { doc.tocs.push({ options, updatedPageNumbers: 0 }); void range; return {}; },
      });
    },
    getSelection() { return rw(doc.selection); },
  };
  return wrap(document, "Document", ["changeTrackingMode"]);
}

export interface InstallOptions {
  /** Supported API sets, e.g. `{ WordApi: "1.3" }`. Defaults to everything the executor uses. */
  sets?: Record<string, string>;
  url?: string;
}

/** Install `Word` and `Office` globals backed by `doc`. Returns the doc for arranging and asserting. */
export function installWordFake(doc: FakeWordDoc = new FakeWordDoc(["Intro", "Body text"]), options: InstallOptions = {}): FakeWordDoc {
  if (options.sets) doc.sets = { ...options.sets };
  if (options.url !== undefined) doc.url = options.url;
  const g = globalThis as Record<string, unknown>;
  const document = makeDocument(doc);
  g.Word = {
    run: async (callback: (context: unknown) => Promise<unknown>) => {
      pendingLoads.clear();
      const context = { document, sync: async () => { doc.syncs++; commitLoads(); } };
      return callback(context);
    },
  };
  g.Office = {
    context: {
      requirements: { isSetSupported: (set: string, version: string) => doc.supported(set, version) },
      get document() { return { url: doc.url }; },
    },
  };
  return doc;
}

export function uninstallWordFake(): void {
  const g = globalThis as Record<string, unknown>;
  delete g.Word;
  delete g.Office;
}
