/**
 * A small in-memory stand-in for the parts of Office.js / Excel the executor
 * uses. It applies changes immediately (no batching), enforces the dimension
 * rules real Excel enforces, and records what was formatted. Install it with
 * `installOfficeMock()` before calling the executor.
 */
type Cell = { v: string | number | boolean | null; f: string | null };
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

export class OfficeError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message);
  }
}


/* ----- strict "load before read" semantics ------------------------------------------------------------------
 * Real Office.js proxies throw when a property is read before `load()` + `context.sync()`. The wrappers below
 * enforce that for ranges, sheets and shapes, so executor code that forgets to load is caught by the tests.
 */
const RAW = Symbol("raw");
const pendingReadErrors: Error[] = [];
const pendingLoads = new Set<{ commit(): void }>();

function parseLoadNames(names: unknown): string[] {
  const list = Array.isArray(names) ? names : typeof names === "string" ? names.split(",") : [];
  // `items/x` style names belong to collections, which are not enforced here.
  return list.map((n) => String(n).trim()).filter((n) => n && !n.includes("/"));
}

function strict<T extends object>(target: T, loadable: readonly string[], label: string): T {
  const loaded = new Set<string>();
  const pending = new Set<string>();
  const handle = {
    commit() {
      for (const name of pending) loaded.add(name);
      pending.clear();
    },
  };
  const proxy: T = new Proxy(target, {
    get(t, key) {
      if (key === RAW) return t;
      if (key === "load") {
        return (names?: unknown) => {
          const requested = names === undefined ? [...loadable] : parseLoadNames(names);
          for (const name of requested) {
            if (label === "Shape" && name === "geometricShapeType" && (t as { type?: string }).type === "Line") {
              throw new OfficeError("InvalidOperation", "Cannot load geometricShapeType on a line.");
            }
            pending.add(name);
          }
          pendingLoads.add(handle);
          return proxy;
        };
      }
      if (typeof key === "string" && loadable.includes(key) && !loaded.has(key)) {
        throw new OfficeError(
          "PropertyNotLoaded",
          `The property '${key}' is not available on this ${label}. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`,
        );
      }
      const value = Reflect.get(t, key, t);
      return typeof value === "function" ? value.bind(t) : value;
    },
    set(t, key, value) {
      (t as Record<string | symbol, unknown>)[key] = value;
      return true;
    },
  });
  return proxy;
}

/** The unwrapped object behind a strict proxy (arguments passed between Office objects need no load). */
function raw<T>(value: T): T {
  return ((value as unknown as Record<symbol, T>)?.[RAW] ?? value) as T;
}

const RANGE_PROPS = ["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "values", "formulas", "numberFormat", "left", "top", "width", "height"] as const;
const SHEET_PROPS = ["name", "id", "position", "visibility", "tabColor"] as const;
const SHAPE_PROPS = ["rotation", "id", "isNullObject", "name", "type", "geometricShapeType", "left", "top", "width", "height"] as const;

function makeRange(sheet: MockSheet, rowIndex: number, columnIndex: number, rowCount: number, columnCount: number, isNullObject = false): MockRange {
  return strict(new MockRange(sheet, rowIndex, columnIndex, rowCount, columnCount, isNullObject), RANGE_PROPS, "Range");
}

const sheetProxies = new WeakMap<MockSheet, MockSheet>();
function wrapSheet(sheet: MockSheet): MockSheet {
  const existing = sheetProxies.get(raw(sheet));
  if (existing) return existing;
  const proxy = strict(raw(sheet), SHEET_PROPS, "Worksheet");
  sheetProxies.set(raw(sheet), proxy);
  return proxy;
}

const shapeProxies = new WeakMap<MockShape, MockShape>();
function wrapShape(shape: MockShape): MockShape {
  const existing = shapeProxies.get(raw(shape));
  if (existing) return existing;
  const proxy = strict(raw(shape), SHAPE_PROPS, "Shape");
  shapeProxies.set(raw(shape), proxy);
  return proxy;
}

/** `collection.load("items/a,items/b")` loads those properties on every item. */
function loadItems(items: Array<{ load(names?: unknown): unknown }>, names: unknown): void {
  const itemProps = (Array.isArray(names) ? names : String(names ?? "").split(","))
    .map((n) => String(n).trim())
    .filter((n) => n.startsWith("items/"))
    .map((n) => n.slice("items/".length));
  if (itemProps.length > 0) for (const item of items) item.load(itemProps.join(","));
}

function colToNum(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
function numToCol(n: number): string {
  let out = "";
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function parseAddress(address: string): { r0: number; c0: number; rows: number; cols: number } {
  const a = address.replace(/\$/g, "");
  let m = /^(\d+):(\d+)$/.exec(a);
  if (m) return { r0: +m[1]! - 1, c0: 0, rows: +m[2]! - +m[1]! + 1, cols: MAX_COL };
  m = /^([A-Za-z]+):([A-Za-z]+)$/.exec(a);
  if (m) return { r0: 0, c0: colToNum(m[1]!) - 1, rows: MAX_ROW, cols: colToNum(m[2]!) - colToNum(m[1]!) + 1 };
  m = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(a);
  if (!m) throw new OfficeError("InvalidArgument", `bad address ${address}`);
  const c1 = colToNum(m[1]!) - 1;
  const r1 = +m[2]! - 1;
  const c2 = m[3] ? colToNum(m[3]) - 1 : c1;
  const r2 = m[4] ? +m[4] - 1 : r1;
  return { r0: Math.min(r1, r2), c0: Math.min(c1, c2), rows: Math.abs(r2 - r1) + 1, cols: Math.abs(c2 - c1) + 1 };
}

/** Shift relative A1 references in a formula by (dr, dc), like fill/paste does. */
function shiftFormula(formula: string, dr: number, dc: number): string {
  return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (_m, ca, col, ra, row) => {
    const c = ca ? col : numToCol(colToNum(col) + dc);
    const r = ra ? row : String(+row + dr);
    return `${ca}${c}${ra}${r}`;
  });
}

const FOUR = ["Rectangle", "RoundRectangle", "Diamond", "FlowChartProcess", "FlowChartDecision", "FlowChartTerminator", "FlowChartInputOutput"];
function siteCount(geometric: string): number {
  return FOUR.includes(geometric) ? 4 : ["Ellipse", "FlowChartConnector"].includes(geometric) ? 8 : 6;
}

export class MockEvents {
  readonly handlers = new Set<(args: any) => unknown>();
  add(handler: (args: any) => unknown) {
    this.handlers.add(handler);
    return { remove: () => { this.handlers.delete(handler); } };
  }
  async emit(args: unknown = {}) { await Promise.all([...this.handlers].map(handler => handler(args))); }
}

let nextShapeId = 0;
export class MockShape {
  readonly onActivated = new MockEvents();
  readonly onDeactivated = new MockEvents();
  readonly id = `shape-${++nextShapeId}`;
  rotation = 0;
  children: MockShape[] = [];
  get group() { return { shapes: { items: this.children.map(wrapShape), load: (names?: unknown) => loadItems(this.children.map(wrapShape), names) } }; }
  isNullObject = false;
  private _box = { left: 0, top: 0, width: 0, height: 0 };
  get left() { return this._box.left; }
  get top() { return this._box.top; }
  get width() { return this._box.width; }
  get height() { return this._box.height; }
  set left(v: number) { this.moveTo({ left: v }); }
  set top(v: number) { this.moveTo({ top: v }); }
  set width(v: number) { this.moveTo({ width: v }); }
  set height(v: number) { this.moveTo({ height: v }); }
  private moveTo(patch: Partial<{ left: number; top: number; width: number; height: number }>) {
    Object.assign(this._box, patch);
  }
  /** Excel raises InvalidOperation for this property on lines/connectors. */
  get geometricShapeType(): string {
    if (this.type === "Line") throw new OfficeError("InvalidOperation", "This operation is not permitted for the current object.");
    return this._geometric;
  }
  set geometricShapeType(value: string) {
    this._geometric = value;
  }
  private _geometric: string;
  name = "";
  deleted = false;
  fillColor: string | null = null;
  /** Mirrors a real Excel quirk: outline weight must be a whole number of points. */
  lineFormat: Record<string, unknown> = new Proxy({} as Record<string, unknown>, {
    set: (target, key, value) => {
      if (key === "weight" && !Number.isInteger(value)) {
        throw new OfficeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
      }
      target[key as string] = value;
      return true;
    },
  });
  text = "";
  textReadFailure: "frame" | "range" | "hasTextSync" | "rangeSync" | null = null;
  emptyTextRangeFails = false;
  textRangeAccesses = 0;
  textFont: Record<string, unknown> = {};
  frame: Record<string, unknown> = {};
  connections: { begin?: { shape: MockShape; site: number }; end?: { shape: MockShape; site: number } } = {};
  lineProps: Record<string, unknown> = {};
  constructor(readonly type: "GeometricShape" | "Line" | "TextBox" | "Group", geometric: string) {
    this._geometric = geometric;
  }
  load() { return this; }
  delete() { this.deleted = true; }
  get fill() {
    // Like Excel, a line has no fill: even navigating to it fails.
    if (this.type === "Line") throw new OfficeError("InvalidOperation", "This operation is not permitted for the current object.");
    return { setSolidColor: (c: string) => { this.fillColor = c; } };
  }
  get textFrame() {
    const self = this;
    // Reading `shape.textFrame` on a line raises InvalidOperation in Excel (seen on a real install).
    if (this.type === "Line" || this.textReadFailure === "frame") throw new OfficeError("InvalidOperation", "This shape has no accessible text frame.");
    return new Proxy(this.frame, {
      get: (t, key) => {
        if (key === "hasText") return Boolean(self.text);
        if (key === "load") return () => {
          if (self.textReadFailure === "hasTextSync") pendingReadErrors.push(new OfficeError("GeneralException", "TextFrame.hasText could not be loaded."));
        };
        if (key === "textRange") {
          self.textRangeAccesses++;
          if (self.textReadFailure === "range" || (self.emptyTextRangeFails && !self.text)) {
            throw new OfficeError("GeneralException", "TextRange._OnAccess: shape has no accessible text story.");
          }
          return {
            load: () => {
              if (self.textReadFailure === "rangeSync") pendingReadErrors.push(new OfficeError("GeneralException", "TextRange._OnAccess failed during sync."));
            },
            get text() { return self.text; }, set text(v: string) { self.text = v; }, font: self.textFont,
          };
        }
        return (t as any)[key];
      },
      set: (t, key, value) => ((t as any)[key] = value, true),
    });
  }
  get line() {
    const self = this;
    const connect = (end: "begin" | "end") => (shapeArg: MockShape, site: number) => {
      const shape = raw(shapeArg);
      if (site < 0 || site >= siteCount(shape.geometricShapeType)) throw new OfficeError("InvalidArgument", `connection site ${site} out of range`);
      self.connections[end] = { shape, site };
    };
    return new Proxy(this.lineProps, {
      get: (t, key) => key === "load" ? () => undefined
        : key === "isBeginConnected" ? Boolean(self.connections.begin)
        : key === "isEndConnected" ? Boolean(self.connections.end)
        : key === "beginConnectedShape" ? self.connections.begin?.shape
        : key === "endConnectedShape" ? self.connections.end?.shape
        : key === "beginConnectedSite" ? self.connections.begin?.site ?? null
        : key === "endConnectedSite" ? self.connections.end?.site ?? null
        : key === "connectorType" ? t[key] ?? "Straight"
        : key === "beginArrowheadStyle" || key === "endArrowheadStyle" ? t[key] ?? "None"
        : key === "connectBeginShape" ? connect("begin") : key === "connectEndShape" ? connect("end") : (t as any)[key],
      set: (t, key, value) => ((t as any)[key] = value, true),
    });
  }
}

export class MockSheet {
  readonly shapeList: MockShape[] = [];
  get shapes() {
    const live = () => this.shapeList.filter((s) => !s.deleted);
    const make = (shape: MockShape) => (this.shapeList.push(shape), wrapShape(shape));
    return {
      get items() { return live().map(wrapShape); },
      load: (names?: unknown) => loadItems(live().map(wrapShape), names),
      getItemOrNullObject: (name: string) => wrapShape(live().find((s) => s.name === name) ?? Object.assign(new MockShape("GeometricShape", ""), { isNullObject: true })),
      addGeometricShape: (type: string) => make(new MockShape("GeometricShape", type)),
      addTextBox: (text: string) => Object.assign(make(new MockShape("TextBox", "TextBox")), { text }),
      addLine: (x1: number, y1: number, x2: number, y2: number, kind: string) => {
        const line = make(new MockShape("Line", ""));
        Object.assign(line, { left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) });
        line.lineProps.connectorType = kind;
        line.lineProps.endpoints = [x1, y1, x2, y2];
        return line;
      },
    };
  }

  readonly cells = new Map<string, Cell>();
  readonly formatLog: Array<{ address: string; props: Record<string, unknown> }> = [];
  readonly noteTexts = new Map<string, string>();
  readonly styles = new Map<string, Record<string, unknown>>();
  readonly merges: Array<{ address: string; across: boolean }> = [];
  readonly sortLog: Array<{ address: string; fields: Array<{ key: number; ascending: boolean }>; hasHeaders: boolean }> = [];
  readonly groupLog: Array<{ op: string; option: string; address: string }> = [];
  readonly outlineLog: Array<[number, number]> = [];
  position = 0;
  tabColor = "";
  autoFilterState = { enabled: false, range: null as string | null, log: [] as unknown[] };
  readonly hiddenRows = new Set<number>();
  readonly hiddenCols = new Set<number>();
  frozen: { rows: number; columns: number } = { rows: 0, columns: 0 };
  visibility = "Visible";
  tables = { items: [] as unknown[], load: () => undefined };

  constructor(
    readonly book: MockWorkbook,
    public name: string,
    readonly id: string,
  ) {}

  get notes() {
    const texts = this.noteTexts;
    const name = this.name;
    return {
      load: () => undefined,
      get items() {
        return [...texts.entries()].map(([address, content]) => ({
          content,
          getLocation: () => ({ address: `${name}!${address}`, load: () => undefined }),
        }));
      },
      getItemOrNullObject: (address: string) => {
        const has = this.noteTexts.has(address);
        const sheet = this;
        return {
          isNullObject: !has,
          load: () => undefined,
          get content() {
            return sheet.noteTexts.get(address) ?? "";
          },
          set content(value: string) {
            sheet.noteTexts.set(address, value);
          },
        };
      },
      add: (address: string, text: string) => {
        if (this.noteTexts.has(address)) throw new OfficeError("InvalidOperation", "note exists");
        this.noteTexts.set(address, text);
      },
    };
  }

  load() {
    return this;
  }
  getRange(address?: string) {
    if (!address) return makeRange(this, 0, 0, MAX_ROW, MAX_COL);
    const g = parseAddress(address);
    return makeRange(this, g.r0, g.c0, g.rows, g.cols);
  }
  showOutlineLevels(rows: number, cols: number) {
    this.outlineLog.push([rows, cols]);
  }
  activate() {
    this.book.active = this.book.sheets.indexOf(this);
  }
  delete() {
    this.book.sheets.splice(this.book.sheets.indexOf(this), 1);
  }
  copy(_where: string, _relative: MockSheet) {
    const clone = this.book.addSheet(`${this.name} (2)`);
    for (const [k, v] of this.cells) clone.cells.set(k, { ...v });
    return wrapSheet(clone);
  }
  get autoFilter() {
    const state = this.autoFilterState;
    return {
      get enabled() {
        return state.enabled;
      },
      load: () => undefined,
      getRangeOrNullObject: () => ({ isNullObject: state.range === null, address: state.range ?? "", load: () => undefined }),
      apply: (rangeArg: MockRange, column?: number, criteria?: unknown) => {
        const range = raw(rangeArg);
        state.enabled = true;
        state.range = range.address;
        state.log.push({ op: "apply", range: range.address.split("!")[1], column, criteria });
      },
      clearCriteria: () => state.log.push({ op: "clear" }),
      reapply: () => state.log.push({ op: "reapply" }),
      remove: () => {
        state.enabled = false;
        state.range = null;
        state.log.push({ op: "remove" });
      },
    };
  }
  getRangeByIndexes(r: number, c: number, rows: number, cols: number) {
    return makeRange(this, r, c, rows, cols);
  }
  getUsedRangeOrNullObject() {
    let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
    for (const [key, cell] of this.cells) {
      if (cell.v === null && cell.f === null) continue;
      const [r, c] = key.split(",").map(Number) as [number, number];
      r0 = Math.min(r0, r); c0 = Math.min(c0, c); r1 = Math.max(r1, r); c1 = Math.max(c1, c);
    }
    if (r1 < 0) return makeRange(this, 0, 0, 1, 1, true);
    return makeRange(this, r0, c0, r1 - r0 + 1, c1 - c0 + 1);
  }
  get freezePanes() {
    return {
      freezeRows: (n: number) => (this.frozen = { rows: n, columns: 0 }),
      freezeColumns: (n: number) => (this.frozen = { rows: 0, columns: n }),
      unfreeze: () => (this.frozen = { rows: 0, columns: 0 }),
    };
  }
  cell(r: number, c: number): Cell {
    return this.cells.get(`${r},${c}`) ?? { v: null, f: null };
  }
  setCell(r: number, c: number, cell: Cell) {
    this.cells.set(`${r},${c}`, cell);
  }
}

export class MockRange {
  constructor(
    readonly sheet: MockSheet,
    readonly rowIndex: number,
    readonly columnIndex: number,
    readonly rowCount: number,
    readonly columnCount: number,
    readonly isNullObject = false,
  ) {}

  /** Default Excel metrics: 48pt wide columns, 15pt tall rows. */
  get left() { return this.columnIndex * 48; }
  get top() { return this.rowIndex * 15; }
  get width() { return this.columnCount * 48; }
  get height() { return this.rowCount * 15; }
  load() {
    return this;
  }
  get address() {
    const name = /[^A-Za-z0-9_]/.test(this.sheet.name) ? `'${this.sheet.name}'` : this.sheet.name;
    const start = `${numToCol(this.columnIndex + 1)}${this.rowIndex + 1}`;
    const end = `${numToCol(this.columnIndex + this.columnCount)}${this.rowIndex + this.rowCount}`;
    return `${name}!${start === end ? start : `${start}:${end}`}`;
  }
  private grid<T>(pick: (cell: Cell) => T): T[][] {
    return Array.from({ length: this.rowCount }, (_, r) =>
      Array.from({ length: this.columnCount }, (_, c) => pick(this.sheet.cell(this.rowIndex + r, this.columnIndex + c))),
    );
  }
  get values() {
    return this.grid((cell) => (cell.f !== null ? computeFormula(this.sheet, cell.f) : (cell.v ?? "")));
  }
  set values(grid: unknown[][]) {
    this.assertShape(grid);
    grid.forEach((row, r) =>
      row.forEach((value, c) => {
        if (value === null) return; // null = leave the cell alone, as in Office.js
        const text = typeof value === "string" && value.startsWith("=") ? value : null;
        this.sheet.setCell(this.rowIndex + r, this.columnIndex + c, text ? { v: null, f: text } : { v: value as Cell["v"], f: null });
      }),
    );
  }
  get formulas() {
    return this.grid((cell) => cell.f ?? (cell.v === null ? "" : String(cell.v)));
  }
  set formulas(grid: unknown[][]) {
    this.assertShape(grid);
    grid.forEach((row, r) =>
      row.forEach((value, c) => {
        const text = String(value);
        this.sheet.setCell(
          this.rowIndex + r,
          this.columnIndex + c,
          text.startsWith("=") ? { v: null, f: text } : { v: text === "" ? null : text, f: null },
        );
      }),
    );
  }
  set numberFormat(grid: string[][]) {
    this.assertShape(grid);
    this.style({ numberFormat: grid[0]![0] });
  }
  private assertShape(grid: unknown[][]) {
    if (grid.length !== this.rowCount || grid.some((row) => row.length !== this.columnCount)) {
      throw new OfficeError(
        "InvalidArgument",
        `The number of rows or columns in the input array doesn't match the size or dimensions of the range.`,
      );
    }
  }
  private log(props: Record<string, unknown>) {
    this.sheet.formatLog.push({ address: this.address.split("!")[1]!, props });
  }
  /** Record a format change on the log and on every covered cell (capped, enough for tests). */
  private style(props: Record<string, unknown>) {
    this.log(props);
    if (this.rowCount * this.columnCount > 5_000) return;
    for (let r = 0; r < this.rowCount; r++)
      for (let c = 0; c < this.columnCount; c++) {
        const key = `${this.rowIndex + r},${this.columnIndex + c}`;
        const record = this.sheet.styles.get(key) ?? {};
        for (const [name, value] of Object.entries(props)) record[name] = value;
        this.sheet.styles.set(key, record);
      }
  }
  get format() {
    const log = (props: Record<string, unknown>) => this.style(props);
    const proxy = (prefix: string) =>
      new Proxy({}, { set: (_t, key, value) => (log({ [`${prefix}${String(key)}`]: value }), true) });
    const target: Record<string, unknown> = {
      font: proxy("font."),
      fill: proxy("fill."),
      borders: {
        getItem: (edge: string) => new Proxy({}, { set: (_t, key, value) => (log({ [`border.${edge}.${String(key)}`]: value }), true) }),
      },
      autofitColumns: () => log({ autofitColumns: true }),
      autofitRows: () => log({ autofitRows: true }),
    };
    return new Proxy(target, {
      set: (_t, key, value) => (log({ [String(key)]: value }), true),
    });
  }
  get numberFormat() {
    return this.grid((_cell) => "General").map((row, r) =>
      row.map((_v, c) => (this.sheet.styles.get(`${this.rowIndex + r},${this.columnIndex + c}`)?.["numberFormat"] as string) ?? "General"),
    );
  }
  getCellProperties() {
    const side = (key: string, edge: string, record: Record<string, unknown>) => ({
      style: (record[`border.${edge}.style`] as string) ?? "None",
      color: (record[`border.${edge}.color`] as string) ?? "#000000",
      weight: (record[`border.${edge}.weight`] as string) ?? "Thin",
    });
    const value = Array.from({ length: this.rowCount }, (_, r) =>
      Array.from({ length: this.columnCount }, (_, c) => {
        const rec = this.sheet.styles.get(`${this.rowIndex + r},${this.columnIndex + c}`) ?? {};
        return {
          format: {
            font: {
              bold: (rec["font.bold"] as boolean) ?? false,
              italic: (rec["font.italic"] as boolean) ?? false,
              color: (rec["font.color"] as string) ?? "#000000",
              name: (rec["font.name"] as string) ?? "Calibri",
              size: (rec["font.size"] as number) ?? 11,
              underline: (rec["font.underline"] as string) ?? "None",
              strikethrough: (rec["font.strikethrough"] as boolean) ?? false,
            },
            fill: { color: (rec["fill.color"] as string) ?? "#FFFFFF" },
            horizontalAlignment: (rec["horizontalAlignment"] as string) ?? "General",
            borders: {
              top: side("top", "EdgeTop", rec),
              bottom: side("bottom", "EdgeBottom", rec),
              left: side("left", "EdgeLeft", rec),
              right: side("right", "EdgeRight", rec),
            },
          },
        };
      }),
    );
    return { value };
  }
  getEntireColumn() {
    return makeRange(this.sheet, 0, this.columnIndex, MAX_ROW, this.columnCount);
  }
  getEntireRow() {
    return makeRange(this.sheet, this.rowIndex, 0, this.rowCount, MAX_COL);
  }
  getImage() {
    return { value: "BASE64PNG" };
  }
  merged: string[] = [];
  merge(across: boolean) {
    this.sheet.merges.push({ address: this.address.split("!")[1]!, across });
    for (let r = 0; r < this.rowCount; r++)
      for (let c = 0; c < this.columnCount; c++) {
        const keep = across ? c === 0 : r === 0 && c === 0;
        if (!keep) this.sheet.cells.delete(`${this.rowIndex + r},${this.columnIndex + c}`);
      }
  }
  unmerge() {
    this.sheet.merges.push({ address: this.address.split("!")[1]!, across: false, unmerged: true } as never);
  }
  get sort() {
    return {
      apply: (fields: Array<{ key: number; ascending: boolean }>, _matchCase: boolean, hasHeaders: boolean) => {
        this.sheet.sortLog.push({ address: this.address.split("!")[1]!, fields, hasHeaders });
        const startRow = hasHeaders ? 1 : 0;
        const rows = [];
        for (let r = startRow; r < this.rowCount; r++) {
          rows.push(Array.from({ length: this.columnCount }, (_, c) => this.sheet.cell(this.rowIndex + r, this.columnIndex + c)));
        }
        const cmp = (a: Cell, b: Cell) => (String(a.v ?? "") < String(b.v ?? "") ? -1 : String(a.v ?? "") > String(b.v ?? "") ? 1 : 0);
        const numeric = (x: Cell) => (typeof x.v === "number" ? x.v : Number.NaN);
        rows.sort((ra, rb) => {
          for (const f of fields) {
            const a = ra[f.key]!;
            const b = rb[f.key]!;
            const diff = Number.isNaN(numeric(a)) || Number.isNaN(numeric(b)) ? cmp(a, b) : numeric(a) - numeric(b);
            if (diff !== 0) return f.ascending ? diff : -diff;
          }
          return 0;
        });
        rows.forEach((cells, i) => cells.forEach((cell, c) => this.sheet.setCell(this.rowIndex + startRow + i, this.columnIndex + c, cell)));
      },
    };
  }
  group(option: string) {
    this.sheet.groupLog.push({ op: "group", option, address: this.address.split("!")[1]! });
  }
  ungroup(option: string) {
    this.sheet.groupLog.push({ op: "ungroup", option, address: this.address.split("!")[1]! });
  }
  set rowHidden(hidden: boolean) {
    for (let r = 0; r < this.rowCount; r++) hidden ? this.sheet.hiddenRows.add(this.rowIndex + r + 1) : this.sheet.hiddenRows.delete(this.rowIndex + r + 1);
  }
  set columnHidden(hidden: boolean) {
    for (let c = 0; c < this.columnCount; c++) hidden ? this.sheet.hiddenCols.add(this.columnIndex + c + 1) : this.sheet.hiddenCols.delete(this.columnIndex + c + 1);
  }
  getCell(r: number, c: number) {
    return makeRange(this.sheet, this.rowIndex + r, this.columnIndex + c, 1, 1);
  }
  getResizedRange(dr: number, dc: number) {
    return makeRange(this.sheet, this.rowIndex, this.columnIndex, this.rowCount + dr, this.columnCount + dc);
  }
  getIntersectionOrNullObject(otherRange: MockRange) {
    const other = raw(otherRange);
    const r0 = Math.max(this.rowIndex, other.rowIndex);
    const c0 = Math.max(this.columnIndex, other.columnIndex);
    const r1 = Math.min(this.rowIndex + this.rowCount, other.rowIndex + other.rowCount);
    const c1 = Math.min(this.columnIndex + this.columnCount, other.columnIndex + other.columnCount);
    if (r1 <= r0 || c1 <= c0) return makeRange(this.sheet, 0, 0, 1, 1, true);
    return makeRange(this.sheet, r0, c0, r1 - r0, c1 - c0);
  }
  clear(applyTo: string) {
    for (let r = 0; r < this.rowCount; r++)
      for (let c = 0; c < this.columnCount; c++) {
        if (applyTo === "Contents" || applyTo === "All") this.sheet.cells.delete(`${this.rowIndex + r},${this.columnIndex + c}`);
      }
    if (applyTo === "Formats" || applyTo === "All") this.log({ cleared: applyTo });
  }
  insert(direction: string) {
    this.shift(direction === "Down" ? "rows" : "cols", 1);
  }
  delete(direction: string) {
    this.shift(direction === "Up" ? "rows" : "cols", -1);
  }
  private shift(axis: "rows" | "cols", sign: 1 | -1) {
    const size = axis === "rows" ? this.rowCount : this.columnCount;
    const at = axis === "rows" ? this.rowIndex : this.columnIndex;
    const moved = new Map<string, Cell>();
    for (const [key, cell] of this.sheet.cells) {
      const [r, c] = key.split(",").map(Number) as [number, number];
      const pos = axis === "rows" ? r : c;
      if (sign === -1 && pos >= at && pos < at + size) continue;
      const next = pos >= at ? pos + sign * size : pos;
      moved.set(axis === "rows" ? `${next},${c}` : `${r},${next}`, cell);
    }
    this.sheet.cells.clear();
    for (const [k, v] of moved) this.sheet.cells.set(k, v);
  }
  copyFrom(sourceRange: MockRange, type = "All") {
    const source = raw(sourceRange);
    for (let r = 0; r < this.rowCount; r++)
      for (let c = 0; c < this.columnCount; c++) {
        // Tile the source over the destination; each tile shifts relative references by its own offset.
        const sr = r % source.rowCount;
        const sc = c % source.columnCount;
        const srcRow = source.rowIndex + sr;
        const srcCol = source.columnIndex + sc;
        const dr = this.rowIndex + r - srcRow;
        const dc = this.columnIndex + c - srcCol;
        const src = this.sheet.cell(srcRow, srcCol);
        const dest: [number, number] = [this.rowIndex + r, this.columnIndex + c];
        if (type === "Formats") {
          const style = this.sheet.styles.get(`${srcRow},${srcCol}`);
          if (style) this.sheet.styles.set(`${dest[0]},${dest[1]}`, { ...style });
          continue;
        }
        const value = type === "Values" && src.f ? { v: computeFormula(this.sheet, src.f), f: null } : src.f ? { v: null, f: shiftFormula(src.f, dr, dc) } : { ...src };
        this.sheet.setCell(dest[0], dest[1], value as Cell);
        if (type === "All") {
          const style = this.sheet.styles.get(`${srcRow},${srcCol}`);
          if (style) this.sheet.styles.set(`${dest[0]},${dest[1]}`, { ...style });
        }
      }
  }
}

/** Evaluate only what tests need: a bare `=A1*B1`-style product/sum of cell refs. */
function computeFormula(sheet: MockSheet, formula: string): number | string {
  const expr = formula.slice(1);
  const resolved = expr.replace(/\$?([A-Z]{1,3})\$?(\d+)/g, (_m, col, row) => {
    const v = sheet.cell(+row - 1, colToNum(col) - 1);
    return String(v.f ? computeFormula(sheet, v.f) : (v.v ?? 0));
  });
  if (!/^[\d+\-*/().\s]+$/.test(resolved)) return formula;
  try {
    return Function(`"use strict"; return (${resolved})`)() as number;
  } catch {
    return formula;
  }
}

export class MockWorkbook {
  readonly worksheetActivated = new MockEvents();
  readonly trackedObjects = new Set<object>();
  sheets: MockSheet[] = [];
  active = 0;
  selection = { sheet: "Sheet1", address: "A1" };
  /** Make `formulas` reads fail, to prove the overwrite guard fails closed. */
  failFormulaReads = false;
  supportsNotes = true;
  supportsShapes = true;
  supportsInsertFromFile = true;
  /** Simulates a shape/chart being selected: getSelectedRange throws like Excel does. */
  invalidSelection = false;
  supportsActiveShape = true;
  activeShape: MockShape | null = null;
  readonly inserted: Array<{ base64: string; options: { sheetNamesToInsert?: string[]; positionType: string; relativeTo?: string } }> = [];

  addSheet(name: string): MockSheet {
    const sheet = new MockSheet(this, name, String(this.sheets.length + 1));
    this.sheets.push(sheet);
    return sheet;
  }
  sheet(name: string): MockSheet {
    const sheet = this.sheets.find((s) => s.name === name || s.id === name);
    if (!sheet) throw new OfficeError("ItemNotFound", `The requested resource doesn't exist: ${name}`);
    return sheet;
  }
  fill(sheet: string, rows: Array<Array<string | number | null>>, at = "A1") {
    const g = parseAddress(at);
    rows.forEach((row, r) => row.forEach((v, c) => v !== null && this.sheet(sheet).setCell(g.r0 + r, g.c0 + c, { v, f: null })));
  }
  get(sheet: string, address: string) {
    const g = parseAddress(address);
    const cell = this.sheet(sheet).cell(g.r0, g.c0);
    return cell.f ?? cell.v;
  }
}

let current: MockWorkbook | null = null;
let patched = false;

export function installOfficeMock(book: MockWorkbook): void {
  pendingReadErrors.length = 0;
  const context = {
    trackedObjects: {
      add: (object: object) => { book.trackedObjects.add(object); },
      remove: (object: object) => { book.trackedObjects.delete(object); },
    },
    workbook: {
      worksheets: {
        onActivated: book.worksheetActivated,
        get items() {
          return book.sheets.map(wrapSheet);
        },
        load: (names?: unknown) => loadItems(book.sheets.map(wrapSheet), names),
        getItem: (key: string) => wrapSheet(book.sheet(key)),
        getItemOrNullObject: (key: string) => wrapSheet(book.sheets.find((x) => x.name === key || x.id === key) ?? Object.assign(new MockSheet(book, "", "0"), { isNullObject: true })),
        getActiveWorksheet: () => wrapSheet(book.sheets[book.active]!),
        add: (name?: string) => wrapSheet(book.addSheet(name ?? `Sheet${book.sheets.length + 1}`)),
      },
      insertWorksheetsFromBase64: (base64: string, options: { sheetNamesToInsert?: string[]; positionType: string; relativeTo?: string }) => {
        book.inserted.push({ base64, options });
        for (const name of options.sheetNamesToInsert ?? ["Template Sheet"]) book.addSheet(name);
      },
      getActiveShapeOrNullObject: () => book.activeShape ? wrapShape(book.activeShape) : {
        isNullObject: true,
        load: (names?: unknown) => {
          if (parseLoadNames(names).some(name => name !== "isNullObject")) throw new OfficeError("InvalidReference", "Reference isn't valid: no active shape.");
        },
      },
      getSelectedRange: () => {
        if (book.invalidSelection) throw new OfficeError("InvalidSelection", "The current selection is invalid for this operation.");
        return book.sheet(book.selection.sheet).getRange(book.selection.address);
      },
    },
    sync: async () => {
      for (const handle of pendingLoads) handle.commit();
      pendingLoads.clear();
      const errors = pendingReadErrors.splice(0);
      if (errors.length) throw errors[0];
    },
  };
  current = book;
  if (!patched) {
    patched = true;
    // Reading formulas can be made to fail, to prove the overwrite guard fails closed.
    const original = Object.getOwnPropertyDescriptor(MockRange.prototype, "formulas")!;
    Object.defineProperty(MockRange.prototype, "formulas", {
      ...original,
      get() {
        if (current?.failFormulaReads) throw new OfficeError("GeneralException", "formulas unavailable");
        return original.get!.call(this);
      },
    });
  }

  (globalThis as any).Excel = {
    run: async (ctxOrCallback: typeof context | ((ctx: typeof context) => Promise<unknown>), callback?: (ctx: typeof context) => Promise<unknown>) =>
      (callback ?? (ctxOrCallback as (ctx: typeof context) => Promise<unknown>))(context),
    RangeUnderlineStyle: { single: "Single", none: "None" },
    InsertShiftDirection: { down: "Down", right: "Right" },
    DeleteShiftDirection: { up: "Up", left: "Left" },
    ClearApplyTo: { all: "All", contents: "Contents", formats: "Formats" },
    RangeCopyType: { formulas: "Formulas", all: "All", values: "Values", formats: "Formats" },
    WorksheetPositionType: { after: "After", before: "Before", beginning: "Beginning", end: "End" },
    SheetVisibility: { visible: "Visible", hidden: "Hidden", veryHidden: "VeryHidden" },
    SortOrientation: { rows: "Rows", columns: "Columns" },
    ConnectorType: { straight: "Straight", elbow: "Elbow", curve: "Curve" },
    FilterOn: { values: "Values", custom: "Custom" },
    FilterOperator: { and: "And", or: "Or" },
    GroupOption: { byRows: "ByRows", byColumns: "ByColumns" },
    BorderLineStyle: { none: "None", continuous: "Continuous" },
  };
  (globalThis as any).Office = {
    context: {
      document: { url: "C:\\books\\test.xlsx" },
      requirements: { isSetSupported: (_name: string, version: string) =>
        version === "1.19" ? current!.supportsActiveShape : version === "1.18" ? current!.supportsNotes : version === "1.9" ? current!.supportsShapes : version === "1.13" ? current!.supportsInsertFromFile : true },
    },
  };
}
