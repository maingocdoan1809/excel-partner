import { JSDOM } from "jsdom";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

/**
 * In-memory stand-in for PowerPoint.run and the objects the PowerPoint executor touches, with the same "load, then
 * sync, then read" discipline as real Office.js: reading a property that was not loaded and synced throws, writes and
 * loads run in order at sync(), and one failing command rejects the sync. Slides export to / import from REAL zip files
 * built with fflate, so the OOXML paths run against actual bytes. API sets are gated: calling something above the
 * configured PowerPointApi version records a violation (the executor is supposed to check first) and throws ApiNotFound.
 *
 * Verified only against this model, not against real PowerPoint.
 */

export class OfficeError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
  }
}

const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export interface DShape {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text?: string;
  fontSize?: number | null;
  fontColor?: string | null;
  fontName?: string;
  bold?: boolean;
  fillType?: string;
  fillColor?: string;
  autoSize?: string;
  wordWrap?: boolean;
  margins?: { left: number; top: number; right: number; bottom: number };
  placeholder?: string;
  contained?: string | null;
  table?: string[][];
  altTitle?: string;
  altDescr?: string;
  hyperlink?: { address?: string; screenTip?: string };
  children?: DShape[];
  /** Reading the text frame of this shape fails (a placeholder holding a picture, say). */
  textBroken?: boolean;
  /** XML as imported from a file; dropped when the API changes the shape. */
  rawXml?: string;
  lineColor?: string;
  lineWeight?: number;
  shapeKind?: string;
}

export interface DSlide {
  id: string;
  layoutId: string;
  masterId: string;
  shapes: DShape[];
  background?: string;
  externalLinks: string[];
}

interface DLayout {
  id: string;
  name: string;
  shapes: DShape[];
}

interface DMaster {
  id: string;
  name: string;
  layouts: DLayout[];
}

export function shape(init: Partial<DShape> & { name: string }): DShape {
  const type = init.type ?? (init.placeholder ? "Placeholder" : "TextBox");
  return {
    id: init.id ?? "",
    type,
    left: 50,
    top: 50,
    width: 300,
    height: 60,
    ...(init.text !== undefined || type === "TextBox" || type === "Placeholder" || type === "GeometricShape" ? { text: init.text ?? "" } : {}),
    fontSize: type === "TextBox" || type === "GeometricShape" || type === "Placeholder" ? 18 : undefined,
    fontColor: "#000000",
    wordWrap: true,
    autoSize: "AutoSizeNone",
    margins: { left: 7.2, top: 3.6, right: 7.2, bottom: 3.6 },
    ...init,
  };
}

const TEXT = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"]);

/* ------------------------------------------------------------------------------------------------------------- */

class ClientResult<T> {
  private set = false;
  private v!: T;
  assign(value: T): void {
    this.v = value;
    this.set = true;
  }
  get value(): T {
    if (!this.set) throw new Error("ClientResult.value read before sync()");
    return this.v;
  }
}

function parseNames(names: unknown): string[] {
  if (names === undefined) return [];
  const list = Array.isArray(names) ? names : String(typeof names === "object" ? (names as { select?: string }).select ?? "" : names).split(",");
  return list.map((n) => String(n).trim()).filter(Boolean);
}

class Ctx {
  private commands: Array<() => void> = [];
  syncs = 0;
  constructor(readonly deck: FakeDeck) {}
  enqueue(command: () => void): void {
    this.commands.push(command);
  }
  async sync(): Promise<void> {
    const batch = this.commands;
    this.commands = [];
    this.syncs++;
    for (const command of batch) command();
  }
  readonly presentation = new FakePresentation(this);
}

/** Base for proxies: values are captured when the load command runs at sync(). */
class Obj<D> {
  protected snap = new Map<string, unknown>();
  constructor(
    protected ctx: Ctx,
    protected label: string,
    protected resolve: () => D | undefined,
    protected getters: Record<string, (d: D) => unknown>,
    protected gates: Record<string, number> = {},
    protected orNull = false,
  ) {}

  protected d(): D {
    const value = this.resolve();
    if (value === undefined) throw new OfficeError("ItemNotFound", `${this.label} does not exist`);
    return value;
  }

  load(names?: unknown): this {
    const wanted = names === undefined ? Object.keys(this.getters) : parseNames(names).filter((n) => !n.includes("/"));
    for (const prop of wanted) if (prop !== "isNullObject" && this.gates[prop]) this.ctx.deck.need(this.gates[prop]!, `${this.label}.${prop}`);
    this.ctx.enqueue(() => this.capture(wanted));
    return this;
  }

  /** Snapshot the current values now (used by load commands, which run at sync). */
  capture(wanted: string[]): void {
    for (const prop of wanted) {
      if (prop === "isNullObject") {
        this.snap.set(prop, this.resolve() === undefined);
        continue;
      }
      const getter = this.getters[prop];
      if (!getter) throw new Error(`fake: ${this.label} has no property '${prop}'`);
      if (this.orNull && this.resolve() === undefined) continue;
      this.snap.set(prop, getter(this.d()));
    }
  }

  protected read<T>(prop: string): T {
    if (!this.snap.has(prop)) throw new Error(`${this.label}.${prop} was read before load() + sync()`);
    return this.snap.get(prop) as T;
  }

  protected write(fn: (d: D) => void): void {
    this.ctx.enqueue(() => fn(this.d()));
  }

  get isNullObject(): boolean {
    return this.read("isNullObject");
  }
}

/** Layouts and masters: identity only. */
class FakeNamed<D extends { id: string; name: string }> extends Obj<D> {
  constructor(ctx: Ctx, label: string, resolve: () => D | undefined) {
    super(ctx, label, resolve, { id: (d) => d.id, name: (d) => d.name }, { id: 3, name: 3 });
  }
  get id(): string { return this.read("id"); }
  get name(): string { return this.read("name"); }
}

function itemsOf<T>(collection: { loadedItems?: T[] }, label: string): T[] {
  if (!collection.loadedItems) throw new Error(`${label}.items was read before load() + sync()`);
  return collection.loadedItems;
}

/* ----- shapes ---------------------------------------------------------------------------------------------- */

class FakeFont extends Obj<DShape> {
  constructor(ctx: Ctx, resolve: () => DShape | undefined) {
    super(ctx, "Font", resolve, {
      size: (d) => d.fontSize ?? null,
      color: (d) => d.fontColor ?? null,
      name: (d) => d.fontName ?? "Calibri",
      bold: (d) => d.bold ?? false,
    });
  }
  get size(): number | null { return this.read("size"); }
  get color(): string | null { return this.read("color"); }
  set size(v: number | null) { this.write((d) => { d.fontSize = v; d.rawXml = undefined; }); }
  set color(v: string | null) { this.write((d) => { d.fontColor = v; d.rawXml = undefined; }); }
  set name(v: string | null) { this.write((d) => { d.fontName = v ?? undefined; d.rawXml = undefined; }); }
  set bold(v: boolean | null) { this.write((d) => { d.bold = v ?? undefined; d.rawXml = undefined; }); }
  set italic(_v: boolean | null) { this.write((d) => { d.rawXml = undefined; }); }
}

class FakeTextRange extends Obj<DShape> {
  readonly font: FakeFont;
  readonly paragraphFormat: { horizontalAlignment: string };
  constructor(ctx: Ctx, shapeResolve: () => DShape | undefined) {
    super(ctx, "TextRange", shapeResolve, {
      text: (d) => {
        if (!TEXT.has(d.type) || d.textBroken) throw new OfficeError("InvalidArgument", `shape ${d.name} has no text frame`);
        return (d.text ?? "").replace(/\n/g, "\r");
      },
    });
    this.font = new FakeFont(ctx, shapeResolve);
    const self = this;
    this.paragraphFormat = { set horizontalAlignment(v: string) { self.write((d) => { d.rawXml = undefined; void v; }); } } as { horizontalAlignment: string };
  }
  get text(): string { return this.read("text"); }
  set text(v: string) {
    this.write((d) => {
      if (!TEXT.has(d.type) || d.textBroken) throw new OfficeError("InvalidArgument", `shape ${d.name} has no text frame`);
      d.text = v.replace(/\r\n|\r/g, "\n");
      d.rawXml = undefined;
    });
  }
}

class FakeTextFrame extends Obj<DShape> {
  readonly textRange: FakeTextRange;
  constructor(ctx: Ctx, resolve: () => DShape | undefined) {
    const check = (d: DShape) => {
      if (!TEXT.has(d.type) || d.textBroken) throw new OfficeError("InvalidArgument", `shape ${d.name} has no text frame`);
    };
    super(ctx, "TextFrame", resolve, {
      hasText: (d) => (check(d), (d.text ?? "") !== ""),
      autoSizeSetting: (d) => (check(d), d.autoSize ?? "AutoSizeNone"),
      wordWrap: (d) => (check(d), d.wordWrap ?? true),
      leftMargin: (d) => (check(d), d.margins?.left ?? 7.2),
      topMargin: (d) => (check(d), d.margins?.top ?? 3.6),
      rightMargin: (d) => (check(d), d.margins?.right ?? 7.2),
      bottomMargin: (d) => (check(d), d.margins?.bottom ?? 3.6),
    }, { hasText: 4 });
    this.textRange = new FakeTextRange(ctx, resolve);
  }
  get hasText(): boolean { return this.read("hasText"); }
  get autoSizeSetting(): string { return this.read("autoSizeSetting"); }
  get wordWrap(): boolean { return this.read("wordWrap"); }
  get leftMargin(): number { return this.read("leftMargin"); }
  get topMargin(): number { return this.read("topMargin"); }
  get rightMargin(): number { return this.read("rightMargin"); }
  get bottomMargin(): number { return this.read("bottomMargin"); }
  override load(names?: unknown): this {
    this.ctx.deck.need(4, "textFrame");
    return super.load(names);
  }
}

class FakeFill extends Obj<DShape> {
  constructor(ctx: Ctx, resolve: () => DShape | undefined) {
    super(ctx, "ShapeFill", resolve, { foregroundColor: (d) => d.fillColor ?? "#000000", type: (d) => d.fillType ?? "NoFill" });
  }
  get foregroundColor(): string { return this.read("foregroundColor"); }
  get type(): string { return this.read("type"); }
  setSolidColor(color: string): void { this.write((d) => { d.fillColor = color; d.fillType = "Solid"; d.rawXml = undefined; }); }
}

class FakeTableCell extends Obj<{ table: string[][]; row: number; column: number }> {
  constructor(ctx: Ctx, resolve: () => { table: string[][]; row: number; column: number } | undefined) {
    super(ctx, "TableCell", resolve, { text: (c) => c.table[c.row]![c.column]! }, {}, true);
  }
  get text(): string { return this.read("text"); }
  set text(v: string) { this.write((c) => { c.table[c.row]![c.column] = v; }); }
}

class FakeTable extends Obj<DShape> {
  readonly rows: { add(index?: number | null, count?: number): void; getItemAt(i: number): { index: number }; deleteRows(rows: Array<{ index: number }>): void };
  readonly columns: { add(index?: number | null, count?: number): void; getItemAt(i: number): { index: number }; deleteColumns(cols: Array<{ index: number }>): void };
  constructor(ctx: Ctx, resolve: () => DShape | undefined) {
    super(ctx, "Table", resolve, {
      rowCount: (d) => d.table!.length,
      columnCount: (d) => d.table![0]!.length,
      values: (d) => d.table!.map((row) => [...row]),
    });
    const t = () => this.d();
    this.rows = {
      add: (index, count = 1) => {
        ctx.deck.need(9, "table.rows.add");
        this.write(() => {
          const table = t().table!;
          const at = index === undefined || index === null || index < 0 || index > table.length ? table.length : index;
          for (let i = 0; i < Math.max(1, count); i++) table.splice(at, 0, table[0]!.map(() => ""));
        });
      },
      getItemAt: (index) => ({ index }),
      deleteRows: (rows) => {
        ctx.deck.need(9, "table.rows.deleteRows");
        this.write(() => {
          const table = t().table!;
          const doomed = new Set(rows.map((r) => r.index));
          t().table = table.filter((_, i) => !doomed.has(i));
        });
      },
    };
    this.columns = {
      add: (index, count = 1) => {
        ctx.deck.need(9, "table.columns.add");
        this.write(() => {
          const table = t().table!;
          const at = index === undefined || index === null || index < 0 || index > table[0]!.length ? table[0]!.length : index;
          for (const row of table) row.splice(at, 0, ...Array(Math.max(1, count)).fill(""));
        });
      },
      getItemAt: (index) => ({ index }),
      deleteColumns: (cols) => {
        ctx.deck.need(9, "table.columns.deleteColumns");
        this.write(() => {
          const doomed = new Set(cols.map((c) => c.index));
          t().table = t().table!.map((row) => row.filter((_, i) => !doomed.has(i)));
        });
      },
    };
  }
  get rowCount(): number { return this.read("rowCount"); }
  get columnCount(): number { return this.read("columnCount"); }
  get values(): string[][] { return this.read("values"); }
  getCellOrNullObject(row: number, column: number): FakeTableCell {
    return new FakeTableCell(this.ctx, () => {
      const table = this.resolve()?.table;
      return table && table[row]?.[column] !== undefined ? { table, row, column } : undefined;
    });
  }
}

const SHAPE_GATES: Record<string, number> = { left: 4, top: 4, width: 4, height: 4, name: 4, type: 4, zOrderPosition: 8, altTextTitle: 10, altTextDescription: 10 };

class FakeShape extends Obj<DShape> {
  readonly textFrame: FakeTextFrame;
  readonly fill: FakeFill;
  readonly lineFormat: { color: string; weight: number; visible: boolean };
  readonly placeholderFormat: Obj<DShape> & { type: string; containedType: string | null };
  readonly group: { shapes: FakeShapeCollection; ungroup(): void };
  constructor(ctx: Ctx, resolve: () => DShape | undefined, private owner: () => DSlide | undefined, orNull = false, private ownerSlideObj?: () => FakeSlide) {
    super(ctx, "Shape", resolve, {
      id: (d) => d.id,
      name: (d) => d.name,
      type: (d) => d.type,
      left: (d) => d.left,
      top: (d) => d.top,
      width: (d) => d.width,
      height: (d) => d.height,
      zOrderPosition: (d) => (owner()!.shapes.indexOf(d) + 1),
      altTextTitle: (d) => d.altTitle ?? "",
      altTextDescription: (d) => d.altDescr ?? "",
    }, SHAPE_GATES, orNull);
    this.textFrame = new FakeTextFrame(ctx, resolve);
    this.fill = new FakeFill(ctx, resolve);
    const self = this;
    this.lineFormat = {
      set color(v: string) { self.write((d) => { d.lineColor = v; d.rawXml = undefined; }); },
      set weight(v: number) { self.write((d) => { d.lineWeight = v; d.rawXml = undefined; }); },
      set visible(_v: boolean) { self.write((d) => { d.rawXml = undefined; }); },
    } as { color: string; weight: number; visible: boolean };
    const pf = new Obj<DShape>(ctx, "PlaceholderFormat", resolve, { type: (d) => d.placeholder ?? "Unsupported", containedType: (d) => d.contained ?? null }, { type: 8, containedType: 8 }) as Obj<DShape> & { type: string; containedType: string | null };
    Object.defineProperty(pf, "type", { get: () => (pf as unknown as { read(p: string): string }).read("type") });
    Object.defineProperty(pf, "containedType", { get: () => (pf as unknown as { read(p: string): string | null }).read("containedType") });
    this.placeholderFormat = pf;
    this.group = {
      shapes: new FakeShapeCollection(ctx, () => this.resolve()?.children, () => undefined, 8, true),
      ungroup: () => {
        ctx.deck.need(8, "group.ungroup");
        this.write((d) => {
          const slide = this.owner()!;
          const at = slide.shapes.indexOf(d);
          slide.shapes.splice(at, 1, ...(d.children ?? []));
        });
      },
    };
  }
  get id(): string { return this.read("id"); }
  get name(): string { return this.read("name"); }
  get type(): string { return this.read("type"); }
  get left(): number { return this.read("left"); }
  get top(): number { return this.read("top"); }
  get width(): number { return this.read("width"); }
  get height(): number { return this.read("height"); }
  get zOrderPosition(): number { return this.read("zOrderPosition"); }
  set name(v: string) { this.write((d) => { d.name = v; d.rawXml = undefined; }); }
  set left(v: number) { this.write((d) => { d.left = v; d.rawXml = undefined; }); }
  set top(v: number) { this.write((d) => { d.top = v; d.rawXml = undefined; }); }
  set width(v: number) { this.write((d) => { d.width = v; d.rawXml = undefined; }); }
  set height(v: number) { this.write((d) => { d.height = v; d.rawXml = undefined; }); }
  set altTextTitle(v: string) { this.ctx.deck.need(10, "altTextTitle"); this.write((d) => { d.altTitle = v; d.rawXml = undefined; }); }
  set altTextDescription(v: string) { this.ctx.deck.need(10, "altTextDescription"); this.write((d) => { d.altDescr = v; d.rawXml = undefined; }); }
  delete(): void {
    this.write((d) => {
      const slide = this.owner()!;
      slide.shapes.splice(slide.shapes.indexOf(d), 1);
      this.ctx.deck.log.push(`shape.delete ${d.id}`);
    });
  }
  setZOrder(position: string): void {
    this.ctx.deck.need(8, "setZOrder");
    this.write((d) => {
      const list = this.owner()!.shapes;
      const i = list.indexOf(d);
      list.splice(i, 1);
      const to = { BringToFront: list.length, SendToBack: 0, BringForward: Math.min(list.length, i + 1), SendBackward: Math.max(0, i - 1) }[position];
      list.splice(to ?? i, 0, d);
    });
  }
  getTable(): FakeTable {
    this.ctx.deck.need(8, "getTable");
    return new FakeTable(this.ctx, () => {
      const d = this.resolve();
      if (d && !d.table) throw new OfficeError("InvalidArgument", "shape is not a table");
      return d;
    });
  }
  setHyperlink(options: { address?: string; screenTip?: string } = {}): object {
    this.ctx.deck.need(10, "setHyperlink");
    this.write((d) => { d.hyperlink = options; d.rawXml = undefined; });
    return {};
  }
  getParentSlide(): FakeSlide {
    this.ctx.deck.need(5, "getParentSlide");
    return this.ownerSlideObj!();
  }
}

class FakeShapeCollection {
  loadedItems?: FakeShape[];
  private cache = new Map<DShape, FakeShape>();
  constructor(
    private ctx: Ctx,
    private list: () => DShape[] | undefined,
    private slide: () => DSlide | undefined,
    private gate = 3,
    private scoped = false,
    private slideObj?: () => FakeSlide,
    private slideOf?: (d: DShape) => DSlide | undefined,
  ) {}

  private wrap(d: DShape): FakeShape {
    let w = this.cache.get(d);
    if (!w) {
      const owner = () => this.slide() ?? this.slideOf?.(d) ?? { id: "", layoutId: "", masterId: "", externalLinks: [], shapes: this.list() ?? [] };
      const slideObj = () => (this.slideOf ? this.ctx.deck.slideProxy(this.ctx, this.slideOf(d)!) : this.slideObj!());
      w = new FakeShape(this.ctx, () => this.list()?.find((s) => s.id === d.id), owner, false, slideObj);
      this.cache.set(d, w);
    }
    return w;
  }

  get items(): FakeShape[] {
    return itemsOf(this, "ShapeCollection");
  }

  load(names?: unknown): this {
    this.ctx.deck.need(this.gate, "shapes.load");
    const props = parseNames(names).filter((n) => n.startsWith("items/")).map((n) => n.slice(6));
    for (const prop of props) if (SHAPE_GATES[prop]) this.ctx.deck.need(SHAPE_GATES[prop]!, `shapes.items/${prop}`);
    this.ctx.enqueue(() => {
      const list = this.list();
      if (!list) throw new OfficeError("ItemNotFound", "the slide does not exist");
      this.loadedItems = list.map((d) => this.wrap(d));
      for (const w of this.loadedItems) w.capture(props);
    });
    return this;
  }

  getItem(id: string): FakeShape {
    return new FakeShape(this.ctx, () => this.list()?.find((s) => s.id === id), () => this.slide() ?? { id: "", layoutId: "", masterId: "", externalLinks: [], shapes: this.list() ?? [] }, false, this.slideObj);
  }
  getItemOrNullObject(id: string): FakeShape {
    return new FakeShape(this.ctx, () => this.list()?.find((s) => s.id === id), () => this.slide(), true, this.slideObj);
  }
  getItemAt(index: number): FakeShape {
    return new FakeShape(this.ctx, () => this.list()?.[index], () => this.slide(), false, this.slideObj);
  }

  private create(init: Partial<DShape> & { name: string }, defaults: Partial<DShape> = {}): FakeShape {
    let created: DShape | undefined;
    const proxy = new FakeShape(this.ctx, () => created, () => this.slide(), false, this.slideObj);
    this.ctx.enqueue(() => {
      const slide = this.slide();
      if (!slide) throw new OfficeError("ItemNotFound", "the slide does not exist");
      const next = Math.max(1, ...slide.shapes.map((s) => Number(s.id) || 0)) + 1;
      created = shape({ ...defaults, ...init, id: String(next) });
      slide.shapes.push(created);
    });
    return proxy;
  }

  addTextBox(text: string, options: Partial<DShape> = {}): FakeShape {
    this.ctx.deck.need(4, "addTextBox");
    return this.create({ name: "TextBox", type: "TextBox", text, ...pickGeometry(options) });
  }
  addGeometricShape(kind: string, options: Partial<DShape> = {}): FakeShape {
    this.ctx.deck.need(4, "addGeometricShape");
    return this.create({ name: kind, type: "GeometricShape", text: "", shapeKind: kind, fillType: "Solid", fillColor: "#4472C4", ...pickGeometry(options) });
  }
  addLine(kind = "Straight", options: Partial<DShape> = {}): FakeShape {
    this.ctx.deck.need(4, "addLine");
    return this.create({ name: "Line", type: "Line", shapeKind: kind, ...pickGeometry(options) });
  }
  addTable(rows: number, columns: number, options: Partial<DShape> & { values?: string[][] } = {}): FakeShape {
    this.ctx.deck.need(8, "addTable");
    const values = options.values ?? Array.from({ length: rows }, () => Array(columns).fill(""));
    return this.create({ name: "Table", type: "Table", table: values.map((r) => [...r]), left: options.left ?? 40, top: options.top ?? 100, width: options.width ?? 500, height: options.height ?? rows * 30 });
  }
  addGroup(values: Array<string | FakeShape>): FakeShape {
    this.ctx.deck.need(8, "addGroup");
    let created: DShape | undefined;
    const proxy = new FakeShape(this.ctx, () => created, () => this.slide(), false, this.slideObj);
    this.ctx.enqueue(() => {
      const slide = this.slide()!;
      const ids = values.map((v) => (typeof v === "string" ? v : (v as unknown as { snap: Map<string, string> }).snap.get("id")!));
      const members = ids.map((id) => {
        const found = slide.shapes.find((s) => s.id === id);
        if (!found) throw new OfficeError("ItemNotFound", `shape ${id} does not exist`);
        return found;
      });
      const next = Math.max(...slide.shapes.map((s) => Number(s.id) || 0)) + 1;
      const left = Math.min(...members.map((m) => m.left));
      const top = Math.min(...members.map((m) => m.top));
      created = shape({ id: String(next), name: "Group", type: "Group", left, top, width: Math.max(...members.map((m) => m.left + m.width)) - left, height: Math.max(...members.map((m) => m.top + m.height)) - top, children: members });
      slide.shapes = slide.shapes.filter((s) => !members.includes(s));
      slide.shapes.push(created);
    });
    return proxy;
  }
  getCount(): ClientResult<number> {
    const result = new ClientResult<number>();
    this.ctx.enqueue(() => result.assign(this.list()?.length ?? 0));
    return result;
  }
  get scopedFlag(): boolean { return this.scoped; }
}

function pickGeometry(o: Partial<DShape>): Partial<DShape> {
  const out: Partial<DShape> = {};
  for (const key of ["left", "top", "width", "height"] as const) if (o[key] !== undefined) out[key] = o[key];
  return out;
}

/* ----- slides ---------------------------------------------------------------------------------------------- */

class FakeSlide extends Obj<DSlide> {
  readonly shapes: FakeShapeCollection;
  readonly layout: FakeNamed<DLayout>;
  readonly slideMaster: FakeNamed<DMaster>;
  readonly background: { fill: { setSolidFill(options: { color?: string }): void } };
  constructor(ctx: Ctx, resolve: () => DSlide | undefined, orNull = false) {
    super(ctx, "Slide", resolve, { id: (s) => s.id, index: (s) => ctx.deck.slides.indexOf(s) }, { index: 8 }, orNull);
    this.shapes = new FakeShapeCollection(ctx, () => this.resolve()?.shapes, () => this.resolve(), 3, false, () => this);
    this.layout = new FakeNamed<DLayout>(ctx, "Layout", () => ctx.deck.layoutOf(this.resolve()));
    this.slideMaster = new FakeNamed<DMaster>(ctx, "Master", () => ctx.deck.masterOf(this.resolve()));
    this.background = {
      fill: {
        setSolidFill: (options) => {
          ctx.deck.need(10, "background.fill.setSolidFill");
          this.write((s) => { s.background = options.color; });
        },
      },
    };
  }
  get id(): string { return this.read("id"); }
  get index(): number { return this.read("index"); }
  delete(): void {
    this.ctx.deck.need(2, "slide.delete");
    this.write((s) => {
      if (this.ctx.deck.failSlideDelete) throw new OfficeError("GeneralException", "delete failed");
      this.ctx.deck.log.push(`slide.delete ${s.id}`);
      this.ctx.deck.slides.splice(this.ctx.deck.slides.indexOf(s), 1);
    });
  }
  moveTo(index: number): void {
    this.ctx.deck.need(8, "moveTo");
    this.write((s) => {
      const list = this.ctx.deck.slides;
      if (index < 0 || index >= list.length) throw new OfficeError("InvalidArgument", "slide index out of range");
      list.splice(list.indexOf(s), 1);
      list.splice(index, 0, s);
      this.ctx.deck.log.push(`slide.moveTo ${s.id} ${index}`);
    });
  }
  exportAsBase64(): ClientResult<string> {
    this.ctx.deck.need(8, "exportAsBase64");
    const result = new ClientResult<string>();
    this.ctx.enqueue(() => {
      this.ctx.deck.log.push(`slide.export ${this.d().id}`);
      result.assign(this.ctx.deck.exportSlide(this.d()));
    });
    return result;
  }
  getImageAsBase64(options: { width?: number } = {}): ClientResult<string> {
    this.ctx.deck.need(8, "getImageAsBase64");
    const result = new ClientResult<string>();
    this.ctx.enqueue(() => {
      this.d();
      this.ctx.deck.imageRequests.push(options.width);
      result.assign(this.ctx.deck.imageData);
    });
    return result;
  }
}

class FakeSlideCollection {
  loadedItems?: FakeSlide[];
  constructor(private ctx: Ctx, private scoped?: () => DSlide[]) {}
  get items(): FakeSlide[] {
    return itemsOf(this, "SlideCollection");
  }
  private list(): DSlide[] {
    return this.scoped ? this.scoped() : this.ctx.deck.slides;
  }
  load(names?: unknown): this {
    const props = parseNames(names).filter((n) => n.startsWith("items/")).map((n) => n.slice(6));
    this.ctx.enqueue(() => {
      this.loadedItems = this.list().map((s) => this.ctx.deck.slideProxy(this.ctx, s));
      for (const w of this.loadedItems) w.capture(props);
    });
    return this;
  }
  getItem(id: string): FakeSlide {
    return new FakeSlide(this.ctx, () => this.list().find((s) => s.id === id));
  }
  getItemOrNullObject(id: string): FakeSlide {
    return new FakeSlide(this.ctx, () => this.list().find((s) => s.id === id), true);
  }
  getItemAt(index: number): FakeSlide {
    return new FakeSlide(this.ctx, () => this.list()[index]);
  }
  getCount(): ClientResult<number> {
    const result = new ClientResult<number>();
    this.ctx.enqueue(() => result.assign(this.list().length));
    return result;
  }
  add(options: { slideMasterId?: string; layoutId?: string } = {}): void {
    this.ctx.deck.need(3, "slides.add");
    this.ctx.enqueue(() => this.ctx.deck.addSlide(options));
  }
}

class FakeLayoutCollection {
  loadedItems?: Array<FakeNamed<DLayout> & { shapes: FakeLayoutShapes }>;
  constructor(private ctx: Ctx, private master: () => DMaster) {}
  get items() { return itemsOf(this, "LayoutCollection"); }
  load(names?: unknown): this {
    const props = parseNames(names).filter((n) => n.startsWith("items/")).map((n) => n.slice(6));
    this.ctx.enqueue(() => {
      this.loadedItems = this.master().layouts.map((layout) => {
        const w = new FakeNamed<DLayout>(this.ctx, "Layout", () => layout) as FakeNamed<DLayout> & { shapes: FakeLayoutShapes };
        w.shapes = new FakeLayoutShapes(this.ctx, layout);
        w.capture(props);
        return w;
      });
    });
    return this;
  }
}

class FakeLayoutShapes {
  loadedItems?: Array<{ name: string; type: string }>;
  constructor(private ctx: Ctx, private layout: DLayout) {}
  get items() { return itemsOf(this, "LayoutShapes"); }
  load(): this {
    this.ctx.enqueue(() => { this.loadedItems = this.layout.shapes.map((s) => ({ name: s.name, type: s.type })); });
    return this;
  }
}

class FakeMasterCollection {
  loadedItems?: Array<FakeNamed<DMaster> & { layouts: FakeLayoutCollection }>;
  constructor(private ctx: Ctx) {}
  get items() { return itemsOf(this, "MasterCollection"); }
  load(names?: unknown): this {
    this.ctx.deck.need(3, "slideMasters");
    const props = parseNames(names).filter((n) => n.startsWith("items/")).map((n) => n.slice(6));
    this.ctx.enqueue(() => {
      this.loadedItems = this.ctx.deck.masters.map((master) => {
        const w = new FakeNamed<DMaster>(this.ctx, "Master", () => master) as FakeNamed<DMaster> & { layouts: FakeLayoutCollection };
        w.layouts = new FakeLayoutCollection(this.ctx, () => master);
        w.capture(props);
        return w;
      });
    });
    return this;
  }
}

class FakePresentation {
  readonly slides: FakeSlideCollection;
  readonly slideMasters: FakeMasterCollection;
  readonly pageSetup: Obj<FakeDeck>;
  private snap = new Map<string, unknown>();
  constructor(private ctx: Ctx) {
    this.slides = new FakeSlideCollection(ctx);
    this.slideMasters = new FakeMasterCollection(ctx);
    this.pageSetup = new Obj<FakeDeck>(ctx, "PageSetup", () => ctx.deck, { slideWidth: (d) => d.size.width, slideHeight: (d) => d.size.height }, { slideWidth: 10, slideHeight: 10 });
    Object.defineProperty(this.pageSetup, "slideWidth", { get: () => (this.pageSetup as unknown as { read(p: string): number }).read("slideWidth") });
    Object.defineProperty(this.pageSetup, "slideHeight", { get: () => (this.pageSetup as unknown as { read(p: string): number }).read("slideHeight") });
  }
  get title(): string {
    if (!this.snap.has("title")) throw new Error("presentation.title was read before load() + sync()");
    return this.snap.get("title") as string;
  }
  load(names?: unknown): this {
    const wanted = parseNames(names);
    this.ctx.enqueue(() => { if (wanted.includes("title")) this.snap.set("title", this.ctx.deck.title); });
    return this;
  }
  getSelectedSlides(): FakeSlideCollection {
    this.ctx.deck.need(5, "getSelectedSlides");
    return new FakeSlideCollection(this.ctx, () => this.ctx.deck.selection.slideIds.map((id) => this.ctx.deck.slides.find((s) => s.id === id)!).filter(Boolean));
  }
  getSelectedShapes(): FakeShapeCollection {
    this.ctx.deck.need(5, "getSelectedShapes");
    const deck = this.ctx.deck;
    const refs = () => deck.selection.shapes;
    const collection = new FakeShapeCollection(
      this.ctx,
      () => refs().map(([slideId, shapeId]) => deck.slides.find((s) => s.id === slideId)?.shapes.find((s) => s.id === shapeId)).filter(Boolean) as DShape[],
      () => undefined,
      5,
      true,
      undefined,
      (d) => deck.slides.find((slide) => slide.shapes.includes(d)),
    );
    return collection;
  }
  insertSlidesFromBase64(base64: string, options: { formatting?: string; targetSlideId?: string; sourceSlideIds?: string[] } = {}): void {
    this.ctx.deck.need(2, "insertSlidesFromBase64");
    this.ctx.enqueue(() => this.ctx.deck.insertSlides(base64, options));
  }
}

/* ----- the deck ------------------------------------------------------------------------------------------- */

export interface DeckOptions {
  /** Highest PowerPointApi version the fake reports (default 10). */
  api?: number;
  size?: { width: number; height: number };
}

export class FakeDeck {
  slides: DSlide[] = [];
  masters: DMaster[] = [];
  title = "Fake deck";
  url = "https://contoso.sharepoint.com/sites/team/Shared%20Documents/Quarterly%20Review.pptx";
  size = { width: 960, height: 540 };
  api: number;
  selection: { slideIds: string[]; shapes: Array<[string, string]> } = { slideIds: [], shapes: [] };
  /** API calls made above the configured version: the executor must gate before calling. */
  violations: string[] = [];
  log: string[] = [];
  imageRequests: Array<number | undefined> = [];
  imageData = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg==";
  insertCalls = 0;
  /** Behaviours to test failure paths. */
  failInsert = false;
  dropShapesOnInsert = false;
  duplicateMasterOnInsert = false;
  failSlideDelete = false;
  private nextSlide = 256;

  constructor(options: DeckOptions = {}) {
    this.api = options.api ?? 10;
    if (options.size) this.size = options.size;
    this.masters = [
      {
        id: "m1#",
        name: "Office Theme",
        layouts: [
          { id: "l1#", name: "Title Slide", shapes: [shape({ id: "2", name: "Title 1", placeholder: "CenterTitle", top: 150, height: 100 }), shape({ id: "3", name: "Subtitle 2", placeholder: "Subtitle", top: 280 })] },
          { id: "l2#", name: "Title and Content", shapes: [shape({ id: "2", name: "Title 1", placeholder: "Title", top: 20, height: 80 }), shape({ id: "3", name: "Content Placeholder 2", placeholder: "Body", top: 120, height: 350 })] },
          { id: "l3#", name: "Blank", shapes: [] },
        ],
      },
    ];
  }

  need(version: number, what: string): void {
    if (version > this.api) {
      this.violations.push(`${what} (needs 1.${version}, host has 1.${this.api})`);
      throw new OfficeError("ApiNotFound", `${what} is not supported (PowerPointApi 1.${version})`);
    }
  }

  isSetSupported(set: string, version: string): boolean {
    if (set !== "PowerPointApi") return false;
    return Number(version.split(".")[1]) <= this.api;
  }

  layoutOf(slide: DSlide | undefined): DLayout | undefined {
    return slide ? this.masters.flatMap((m) => m.layouts).find((l) => l.id === slide.layoutId) : undefined;
  }
  masterOf(slide: DSlide | undefined): DMaster | undefined {
    return slide ? this.masters.find((m) => m.id === slide.masterId) : undefined;
  }

  slideProxy(ctx: Ctx, slide: DSlide): FakeSlide {
    return new FakeSlide(ctx, () => this.slides.find((s) => s === slide));
  }

  /** Test setup: a slide made of the given shapes (ids are assigned when missing). Returns its id. */
  addSlideWith(shapes: Array<Partial<DShape> & { name: string }>, options: { layoutId?: string; background?: string; externalLinks?: string[]; at?: number } = {}): string {
    const id = `${this.nextSlide++}#`;
    const used = new Set(shapes.map((s) => s.id).filter(Boolean));
    let next = 2;
    const built = shapes.map((init) => {
      let sid = init.id;
      while (!sid) {
        if (!used.has(String(next))) sid = String(next);
        next++;
      }
      used.add(sid);
      return shape({ ...init, id: sid });
    });
    const slide: DSlide = { id, layoutId: options.layoutId ?? "l2#", masterId: "m1#", shapes: built, background: options.background, externalLinks: options.externalLinks ?? [] };
    if (options.at === undefined) this.slides.push(slide);
    else this.slides.splice(options.at, 0, slide);
    return id;
  }

  slideById(id: string): DSlide {
    const slide = this.slides.find((s) => s.id === id);
    if (!slide) throw new Error(`no slide ${id}`);
    return slide;
  }
  ids(): string[] {
    return this.slides.map((s) => s.id);
  }

  addSlide(options: { slideMasterId?: string; layoutId?: string }): void {
    const master = this.masters.find((m) => m.id === options.slideMasterId) ?? this.masters[0]!;
    const layout = master.layouts.find((l) => l.id === options.layoutId) ?? master.layouts[0]!;
    if (options.layoutId && layout.id !== options.layoutId) throw new OfficeError("InvalidArgument", "layout not found");
    const slide: DSlide = {
      id: `${this.nextSlide++}#`,
      layoutId: layout.id,
      masterId: master.id,
      shapes: layout.shapes.map((s) => ({ ...s, text: "", contained: null, rawXml: undefined })),
      externalLinks: [],
    };
    this.slides.push(slide);
    this.log.push(`slides.add ${slide.id}`);
  }

  /* ----- OOXML ----- */

  exportSlide(slide: DSlide): string {
    const layout = this.layoutOf(slide);
    const rels =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${NS_R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      slide.externalLinks.map((url, i) => `<Relationship Id="rId${i + 2}" Type="${NS_R}/hyperlink" Target="${url}" TargetMode="External"/>`).join("") +
      `</Relationships>`;
    const files: Record<string, Uint8Array> = {
      "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`),
      "ppt/presentation.xml": strToU8(
        `<?xml version="1.0"?><p:presentation xmlns:p="${NS_P}" xmlns:a="${NS_A}"><p:sldSz cx="${Math.round(this.size.width * 12700)}" cy="${Math.round(this.size.height * 12700)}"/></p:presentation>`,
      ),
      "ppt/slides/slide1.xml": strToU8(this.slideXml(slide)),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(rels),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(`<?xml version="1.0"?><p:sldLayout xmlns:p="${NS_P}" name="${layout?.name ?? ""}" data-id="${slide.layoutId}" data-master="${slide.masterId}"/>`),
      "ppt/slideMasters/slideMaster1.xml": strToU8(`<?xml version="1.0"?><p:sldMaster xmlns:p="${NS_P}"/>`),
    };
    return Buffer.from(zipSync(files)).toString("base64");
  }

  private shapeXml(s: DShape): string {
    if (s.rawXml) return s.rawXml;
    const esc = (t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const emu = (n: number) => Math.round(n * 12700);
    const xfrm = `<a:xfrm><a:off x="${emu(s.left)}" y="${emu(s.top)}"/><a:ext cx="${emu(s.width)}" cy="${emu(s.height)}"/></a:xfrm>`;
    const descr = `${s.altDescr ? ` descr="${esc(s.altDescr)}"` : ""}${s.altTitle ? ` title="${esc(s.altTitle)}"` : ""}`;
    if (s.type === "Table") {
      const rows = (s.table ?? []).map((row) => `<a:tr h="370840">${row.map((cell) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(cell)}</a:t></a:r></a:p></a:txBody></a:tc>`).join("")}</a:tr>`).join("");
      return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${s.id}" name="${esc(s.name)}"${descr}/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(s.left)}" y="${emu(s.top)}"/><a:ext cx="${emu(s.width)}" cy="${emu(s.height)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid/>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    }
    if (s.type === "Image") {
      return `<p:pic><p:nvPicPr><p:cNvPr id="${s.id}" name="${esc(s.name)}"${descr}/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip/></p:blipFill><p:spPr>${xfrm}</p:spPr></p:pic>`;
    }
    if (s.type === "Line") {
      return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${s.id}" name="${esc(s.name)}"${descr}/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm}</p:spPr></p:cxnSp>`;
    }
    if (s.type === "Group") {
      return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${s.id}" name="${esc(s.name)}"${descr}/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm}</p:grpSpPr>${(s.children ?? []).map((c) => this.shapeXml(c)).join("")}</p:grpSp>`;
    }
    const ph = s.placeholder ? `<p:ph type="${s.placeholder}"/>` : "";
    const paragraphs = (s.text ?? "").split("\n").map((line) => `<a:p><a:r><a:rPr${s.fontSize ? ` sz="${Math.round(s.fontSize * 100)}"` : ""}/><a:t>${esc(line)}</a:t></a:r></a:p>`).join("");
    return `<p:sp><p:nvSpPr><p:cNvPr id="${s.id}" name="${esc(s.name)}"${descr}/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${xfrm}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
  }

  private slideXml(slide: DSlide): string {
    const bg = slide.background ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${slide.background.replace("#", "").toUpperCase()}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : "";
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${NS_A}" xmlns:p="${NS_P}" xmlns:r="${NS_R}"><p:cSld>${bg}<p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      slide.shapes.map((s) => this.shapeXml(s)).join("") +
      `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
    );
  }

  insertSlides(base64: string, options: { formatting?: string; targetSlideId?: string }): void {
    this.insertCalls++;
    if (this.failInsert) throw new OfficeError("GeneralException", "insertSlidesFromBase64 failed");
    const files = unzipSync(new Uint8Array(Buffer.from(base64, "base64")));
    const rels = strFromU8(files["ppt/slides/_rels/slide1.xml.rels"] ?? new Uint8Array());
    if (/TargetMode="External"/i.test(rels)) throw new OfficeError("IllegalAccessError", "external relationships are not allowed");
    let at = 0;
    if (options.targetSlideId !== undefined) {
      const target = this.slides.findIndex((s) => s.id === options.targetSlideId);
      if (target === -1) throw new OfficeError("SlideNotFound", `slide ${options.targetSlideId} not found`);
      at = target + 1;
    }
    const parser = new (new JSDOM("").window.DOMParser)();
    const xml = parser.parseFromString(strFromU8(files["ppt/slides/slide1.xml"]!), "application/xml");
    const layoutXml = parser.parseFromString(strFromU8(files["ppt/slideLayouts/slideLayout1.xml"]!), "application/xml");
    const layoutId = layoutXml.documentElement.getAttribute("data-id") ?? "l2#";
    let masterId = layoutXml.documentElement.getAttribute("data-master") ?? "m1#";
    if (options.formatting === "KeepSourceFormatting" && this.duplicateMasterOnInsert) {
      const source = this.masters.find((m) => m.id === masterId);
      masterId = `m${this.masters.length + 1}#`;
      this.masters.push({ id: masterId, name: `1_${source?.name ?? "Theme"}`, layouts: (source?.layouts ?? []).map((l) => ({ ...l, id: `${l.id}b` })) });
    }
    const serializer = new (new JSDOM("").window.XMLSerializer)();
    const tree = xml.getElementsByTagNameNS(NS_P, "spTree")[0]!;
    const shapes: DShape[] = [];
    const parse = (children: Element[]): DShape[] => {
      const out: DShape[] = [];
      for (const child of children) {
        if (child.namespaceURI !== NS_P || !["sp", "pic", "graphicFrame", "grpSp", "cxnSp"].includes(child.localName)) continue;
        const cNvPr = child.getElementsByTagNameNS(NS_P, "cNvPr")[0]!;
        const off = child.getElementsByTagNameNS(NS_A, "off")[0];
        const ext = child.getElementsByTagNameNS(NS_A, "ext")[0];
        const ph = child.getElementsByTagNameNS(NS_P, "ph")[0];
        const type = { sp: ph ? "Placeholder" : "TextBox", pic: "Image", graphicFrame: "Table", grpSp: "Group", cxnSp: "Line" }[child.localName]!;
        const text = Array.from(child.getElementsByTagNameNS(NS_A, "p")).map((p) => Array.from(p.getElementsByTagNameNS(NS_A, "t")).map((t) => t.textContent ?? "").join("")).join("\n");
        const rPr = child.getElementsByTagNameNS(NS_A, "rPr")[0];
        const s: DShape = shape({
          id: cNvPr.getAttribute("id") ?? "0",
          name: cNvPr.getAttribute("name") ?? "",
          type,
          left: off ? Number(off.getAttribute("x")) / 12700 : 0,
          top: off ? Number(off.getAttribute("y")) / 12700 : 0,
          width: ext ? Number(ext.getAttribute("cx")) / 12700 : 0,
          height: ext ? Number(ext.getAttribute("cy")) / 12700 : 0,
          placeholder: ph?.getAttribute("type") ?? undefined,
          ...(child.localName === "sp" ? { text, fontSize: rPr?.getAttribute("sz") ? Number(rPr.getAttribute("sz")) / 100 : 18 } : {}),
          altDescr: cNvPr.getAttribute("descr") ?? undefined,
          altTitle: cNvPr.getAttribute("title") ?? undefined,
        });
        if (type === "Table") s.table = Array.from(child.getElementsByTagNameNS(NS_A, "tr")).map((tr) => Array.from(tr.getElementsByTagNameNS(NS_A, "tc")).map((tc) => Array.from(tc.getElementsByTagNameNS(NS_A, "t")).map((t) => t.textContent ?? "").join("")));
        if (type === "Group") s.children = parse(Array.from(child.children).filter((c) => c.localName !== "nvGrpSpPr" && c.localName !== "grpSpPr"));
        s.rawXml = serializer.serializeToString(child);
        out.push(s);
      }
      return out;
    };
    shapes.push(...(this.dropShapesOnInsert ? [] : parse(Array.from(tree.children))));
    const bg = xml.getElementsByTagNameNS(NS_A, "srgbClr")[0]?.parentElement?.parentElement?.parentElement?.localName === "bg" ? `#${xml.getElementsByTagNameNS(NS_A, "srgbClr")[0]!.getAttribute("val")}` : undefined;
    const slide: DSlide = { id: `${this.nextSlide++}#`, layoutId, masterId, shapes, background: bg, externalLinks: [] };
    this.slides.splice(at, 0, slide);
    this.log.push(`insert ${slide.id} at ${at}`);
  }
}

/* ----- installation --------------------------------------------------------------------------------------- */

export interface Installed {
  deck: FakeDeck;
  restore(): void;
}

/** Install globalThis.PowerPoint, Office, DOMParser and XMLSerializer for the executor, backed by `deck`. */
export function installFakePowerPoint(deck: FakeDeck = new FakeDeck()): Installed {
  const g = globalThis as Record<string, unknown>;
  const saved = { PowerPoint: g.PowerPoint, Office: g.Office, DOMParser: g.DOMParser, XMLSerializer: g.XMLSerializer };
  const window = new JSDOM("").window;
  g.DOMParser = window.DOMParser;
  g.XMLSerializer = window.XMLSerializer;
  const shapeTypes = ["Rectangle", "RoundRectangle", "Ellipse", "Triangle", "Diamond", "RightArrow", "Chevron", "Star5", "Hexagon", "Cloud"];
  g.PowerPoint = {
    run: async (callback: (context: Ctx) => Promise<unknown>) => {
      const context = new Ctx(deck);
      const result = await callback(context);
      await context.sync();
      return result;
    },
    GeometricShapeType: Object.fromEntries(shapeTypes.map((n) => [n[0]!.toLowerCase() + n.slice(1), n])),
  };
  g.Office = {
    context: {
      requirements: { isSetSupported: (set: string, version: string) => deck.isSetSupported(set, version) },
      document: {
        get url() { return deck.url; },
        getSelectedDataAsync: (coercion: string, callback: (result: { status: string; value?: unknown; error?: { message: string } }) => void) => {
          const first = deck.slides.find((s) => s.id === deck.selection.slideIds[0]);
          queueMicrotask(() =>
            callback(
              coercion === "slideRange" && first
                ? { status: "succeeded", value: { slides: [{ id: Number(first.id.replace(/\D/g, "")), title: "", index: deck.slides.indexOf(first) + 1 }] } }
                : { status: "succeeded", value: { slides: [] } },
            ),
          );
        },
      },
    },
    CoercionType: { SlideRange: "slideRange" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
  };
  return {
    deck,
    restore() {
      for (const [key, value] of Object.entries(saved)) {
        if (value === undefined) delete g[key];
        else g[key] = value;
      }
    },
  };
}
