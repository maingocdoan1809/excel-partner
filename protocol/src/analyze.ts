import type { AnalyzeRangeInput } from "./tools-advanced.js";

/**
 * The arithmetic of `analyze_range`: profile the columns of a table, or filter, group and aggregate its rows, without the
 * rows ever reaching the conversation. Pure (no Office.js): the task pane reads the cells in blocks and feeds them to a
 * `RangeAnalyzer`, which keeps only running totals.
 */

export type Kind = "number" | "text" | "boolean" | "error" | "empty";
type Cell = string | number | boolean | null;

/** What the caller asked for is wrong (an unknown column); the message says what is available. */
export class AnalyzeError extends Error {
  constructor(message: string) { super(message); this.name = "AnalyzeError"; }
}

export interface AnalyzeColumn {
  /** Position within the analysed range, 0-based. */
  index: number;
  /** The column's letters on the sheet. */
  letter: string;
  /** Header text, or the letters when the range has no header row. */
  name: string;
  /** Number format shows a date or time: values are Excel serial numbers. */
  dateLike: boolean;
}

/** Cells examined per distinct-value table before it stops growing: a column of IDs must not use the memory of the pane. */
export const DISTINCT_CAP = 1000;
export const GROUP_CAP = 5000;
export const ROW_CAP = 20_000;
export const MAX_ROW_COLUMNS = 20;

export function columnLetter(index: number): string {
  let letters = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
}

/** A number format that shows a date and/or a time (`yyyy-mm-dd`, `d/m/yy`, `h:mm`), ignoring quoted text and `[...]` parts. */
export function isDateFormat(format: unknown): boolean {
  if (typeof format !== "string") return false;
  const stripped = format.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").replace(/\\./g, "");
  return /[ymdh]/i.test(stripped) && !/^general$/i.test(format.trim());
}

/** `2025-03-31` as the serial number Excel keeps for that day, or null when the text is not such a date. */
export function isoToSerial(text: string): number | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text.trim());
  if (!match) return null;
  const [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3])];
  const time = Date.UTC(year, month - 1, day);
  const check = new Date(time);
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null;
  return time / 86_400_000 + 25_569;
}

function kindOf(value: unknown, type: unknown): Kind {
  switch (type) {
    case "Empty": return "empty";
    case "Integer": case "Double": return "number";
    case "Boolean": return "boolean";
    case "Error": return "error";
    case "String": return value === "" ? "empty" : "text";
  }
  if (value === null || value === undefined || value === "") return "empty";
  if (typeof value === "number") return "number";
  if (typeof value === "boolean") return "boolean";
  return "text";
}

/** Find a column by header text (any case), else by its letters on the sheet. */
export function resolveColumn(columns: readonly AnalyzeColumn[], reference: string): AnalyzeColumn {
  const wanted = reference.trim().toLowerCase();
  const byName = columns.filter((column) => column.name.trim().toLowerCase() === wanted);
  if (byName.length === 1) return byName[0]!;
  if (byName.length > 1) throw new AnalyzeError(`Column '${reference}' is ambiguous: its name is in ${byName.map((column) => column.letter).join(", ")}. Use the column letters.`);
  const byLetter = columns.filter((column) => column.letter.toLowerCase() === wanted);
  if (byLetter.length === 1) return byLetter[0]!;
  throw new AnalyzeError(`No column '${reference}'. Columns: ${columns.map((column) => `${column.name} (${column.letter})`).join(", ")}.`);
}

const clean = (number: number): number => (Number.isFinite(number) ? Number(number.toPrecision(15)) : number);
const text = (value: Cell): string => String(value);

function toNumber(target: unknown): number | null {
  if (typeof target === "number") return target;
  if (typeof target === "string" && target.trim() !== "" && !Number.isNaN(Number(target))) return Number(target);
  return null;
}

/** Order of a cell against a value: negative, zero, positive; null when they cannot be ordered (a number against words). */
function order(kind: Kind, value: Cell, target: unknown): number | null {
  if (kind === "number") { const wanted = toNumber(target); return wanted === null ? null : (value as number) - wanted; }
  if (kind === "text" && typeof target === "string") return text(value).toLowerCase().localeCompare(target.toLowerCase());
  return null;
}

function equals(kind: Kind, value: Cell, target: unknown): boolean {
  if (kind === "empty") return target === "" || target === null;
  if (kind === "number") { const wanted = toNumber(target); return wanted !== null && value === wanted; }
  if (kind === "boolean") return value === target || String(target).toLowerCase() === String(value);
  if (kind === "text") return text(value).trim().toLowerCase() === String(target).trim().toLowerCase();
  return false;
}

type Test = (value: Cell, kind: Kind) => boolean;

function compile(op: string, a: unknown, b: unknown, list: unknown[] | undefined, dateLike: boolean): Test {
  const date = (target: unknown) => (dateLike && typeof target === "string" ? isoToSerial(target) ?? target : target);
  const first = date(a), second = date(b), candidates = (list ?? []).map(date);
  const compare = (test: (difference: number) => boolean): Test => (value, kind) => { const difference = order(kind, value, first); return difference !== null && test(difference); };
  const needle = String(a ?? "").toLowerCase();
  const contains = (value: Cell, kind: Kind) => kind !== "empty" && kind !== "error" && text(value).toLowerCase().includes(needle);
  switch (op) {
    case "eq": return (value, kind) => equals(kind, value, first);
    case "ne": return (value, kind) => !equals(kind, value, first);
    case "gt": return compare((difference) => difference > 0);
    case "gte": return compare((difference) => difference >= 0);
    case "lt": return compare((difference) => difference < 0);
    case "lte": return compare((difference) => difference <= 0);
    case "between": return (value, kind) => { const low = order(kind, value, first), high = order(kind, value, second); return low !== null && high !== null && low >= 0 && high <= 0; };
    case "in": return (value, kind) => candidates.some((candidate) => equals(kind, value, candidate));
    case "contains": return contains;
    case "notContains": return (value, kind) => !contains(value, kind);
    case "startsWith": return (value, kind) => kind !== "empty" && kind !== "error" && text(value).toLowerCase().startsWith(needle);
    case "endsWith": return (value, kind) => kind !== "empty" && kind !== "error" && text(value).toLowerCase().endsWith(needle);
    case "blank": return (_value, kind) => kind === "empty";
    case "notBlank": return (_value, kind) => kind !== "empty";
    default: throw new AnalyzeError(`Unknown condition '${op}'.`);
  }
}

interface Accumulator {
  /** Cells with a value (errors excluded) or, for `count` without a column, rows. */
  count: number;
  numeric: number;
  sum: number;
  min: number | null;
  max: number | null;
  ignored: number;
  distinct: Set<string>;
  capped: boolean;
}
const newAccumulator = (): Accumulator => ({ count: 0, numeric: 0, sum: 0, min: null, max: null, ignored: 0, distinct: new Set(), capped: false });

interface AggregateSpec { fn: "count" | "sum" | "avg" | "min" | "max" | "distinct"; column: AnalyzeColumn | null; name: string }

function feed(accumulator: Accumulator, spec: AggregateSpec, value: Cell, kind: Kind): void {
  if (!spec.column) { accumulator.count++; return; }
  if (kind === "empty") return;
  if (kind === "error") { accumulator.ignored++; return; }
  accumulator.count++;
  if (spec.fn === "distinct") {
    if (accumulator.distinct.size < DISTINCT_CAP) accumulator.distinct.add(`${kind}:${text(value).toLowerCase()}`);
    else if (!accumulator.distinct.has(`${kind}:${text(value).toLowerCase()}`)) accumulator.capped = true;
    return;
  }
  if (spec.fn === "count") return;
  if (kind !== "number") { accumulator.ignored++; return; }
  const number = value as number;
  accumulator.numeric++; accumulator.sum += number;
  accumulator.min = accumulator.min === null ? number : Math.min(accumulator.min, number);
  accumulator.max = accumulator.max === null ? number : Math.max(accumulator.max, number);
}

function resultOf(accumulator: Accumulator, spec: AggregateSpec): Cell {
  switch (spec.fn) {
    case "count": return accumulator.count;
    case "distinct": return accumulator.distinct.size;
    case "sum": return clean(accumulator.sum);
    case "avg": return accumulator.numeric ? clean(accumulator.sum / accumulator.numeric) : null;
    case "min": return accumulator.min;
    case "max": return accumulator.max;
  }
}

interface Group { key: Cell[]; rows: number; accumulators: Accumulator[] }

interface ColumnProfile {
  nonEmpty: number; blank: number; numbers: number; texts: number; booleans: number; errors: number;
  sum: number; min: number | null; max: number | null;
  values: Map<string, { value: Cell; count: number }>; capped: boolean; sample: Cell[];
}
const newProfile = (): ColumnProfile => ({ nonEmpty: 0, blank: 0, numbers: 0, texts: 0, booleans: 0, errors: 0, sum: 0, min: null, max: null, values: new Map(), capped: false, sample: [] });

export class RangeAnalyzer {
  readonly notes: string[] = [];
  private scanned = 0;
  private matched = 0;
  private readonly filters: Array<{ column: AnalyzeColumn; test: Test }> = [];
  private readonly groupColumns: AnalyzeColumn[];
  private readonly specs: AggregateSpec[];
  private readonly grouping: boolean;
  private readonly listing: boolean;
  private readonly selected: AnalyzeColumn[];
  private readonly listOrder: AnalyzeColumn | null;
  private readonly groups = new Map<string, Group>();
  private groupsCapped = false;
  private readonly totals: Group;
  private readonly listed: Array<{ row: number; values: Cell[]; sort: Cell }> = [];
  private readonly profiled: AnalyzeColumn[];
  private readonly profiles = new Map<number, ColumnProfile>();
  private errorCells = 0;

  constructor(private readonly columns: readonly AnalyzeColumn[], private readonly input: AnalyzeRangeInput) {
    this.profiled = input.columns?.length ? input.columns.map((reference) => resolveColumn(columns, reference)) : columns.slice(0, 60) as AnalyzeColumn[];
    for (const column of this.profiled) this.profiles.set(column.index, newProfile());
    if (input.mode === "profile" && !input.columns?.length && columns.length > 60) this.notes.push(`Only the first 60 of ${columns.length} columns are profiled; name the others in \`columns\`.`);

    for (const filter of input.where) {
      const column = resolveColumn(columns, filter.column);
      this.filters.push({ column, test: compile(filter.op, filter.value, filter.value2, filter.values, column.dateLike) });
    }
    this.groupColumns = input.groupBy.map((reference) => resolveColumn(columns, reference));
    const taken = new Set<string>();
    const requested = input.aggregate.map((aggregate): AggregateSpec => {
      const column = aggregate.column === undefined ? null : resolveColumn(columns, aggregate.column);
      const wanted = aggregate.as ?? (column ? `${aggregate.fn}_${column.name}` : "count");
      let name = wanted;
      for (let suffix = 2; taken.has(name); suffix++) name = `${wanted}_${suffix}`;
      taken.add(name);
      return { fn: aggregate.fn, column, name };
    });
    this.grouping = this.groupColumns.length > 0;
    this.specs = requested.length || !this.grouping ? requested : [{ fn: "count", column: null, name: "count" }];
    this.listing = !this.grouping && this.specs.length === 0;
    this.listOrder = this.listing && input.orderBy ? resolveColumn(columns, input.orderBy.by) : null;
    const wanted = input.columns?.length ? this.profiled : columns;
    this.selected = wanted.slice(0, MAX_ROW_COLUMNS) as AnalyzeColumn[];
    if (input.mode === "query" && this.listing && wanted.length > MAX_ROW_COLUMNS) this.notes.push(`Only the first ${MAX_ROW_COLUMNS} columns are listed; name the ones you need in \`columns\`.`);
    this.totals = { key: [], rows: 0, accumulators: this.specs.map(newAccumulator) };
    const dates = columns.filter((column) => column.dateLike).map((column) => column.name);
    if (dates.length) this.notes.push(`${dates.join(", ")} hold dates or times as Excel serial numbers (45000 is 2023-03-15). Filter them with YYYY-MM-DD.`);
  }

  /** Feed a block of rows. `firstRow` is the sheet row number of the block's first row. */
  add(values: readonly (readonly unknown[])[], types: readonly (readonly unknown[])[] | undefined, firstRow: number): void {
    for (let r = 0; r < values.length; r++) {
      const row = values[r]!, typeRow = types?.[r] ?? [];
      const kinds = row.map((value, c) => kindOf(value, typeRow[c]));
      if (kinds.every((kind) => kind === "empty")) continue;
      this.scanned++;
      if (this.input.mode === "profile") { this.profile(row, kinds); continue; }
      let keep = true;
      for (const { column, test } of this.filters) if (!test(row[column.index] as Cell, kinds[column.index]!)) { keep = false; break; }
      if (!keep) continue;
      this.matched++;
      kinds.forEach((kind) => { if (kind === "error") this.errorCells++; });
      if (this.listing) this.list(row as Cell[], kinds, firstRow + r);
      else this.collect(row as Cell[], kinds);
    }
  }

  private profile(row: readonly unknown[], kinds: readonly Kind[]): void {
    for (const column of this.profiled) {
      const profile = this.profiles.get(column.index)!, kind = kinds[column.index]!, value = row[column.index] as Cell;
      if (kind === "empty") { profile.blank++; continue; }
      if (kind === "error") { profile.errors++; continue; }
      profile.nonEmpty++;
      if (kind === "number") { profile.numbers++; const n = value as number; profile.sum += n; profile.min = profile.min === null ? n : Math.min(profile.min, n); profile.max = profile.max === null ? n : Math.max(profile.max, n); }
      else if (kind === "text") profile.texts++;
      else profile.booleans++;
      const key = `${kind}:${text(value).toLowerCase()}`, known = profile.values.get(key);
      if (known) known.count++;
      else if (profile.values.size < DISTINCT_CAP) profile.values.set(key, { value, count: 1 });
      else profile.capped = true;
      if (profile.sample.length < 3 && !profile.sample.includes(value)) profile.sample.push(value);
    }
  }

  private collect(row: Cell[], kinds: readonly Kind[]): void {
    let group: Group;
    if (this.grouping) {
      const key = this.groupColumns.map((column) => (kinds[column.index] === "empty" ? null : row[column.index]!));
      const id = JSON.stringify(key.map((part) => (typeof part === "string" ? part.toLowerCase() : part)));
      const known = this.groups.get(id);
      if (known) group = known;
      else if (this.groups.size >= GROUP_CAP) { this.groupsCapped = true; return; }
      else { group = { key, rows: 0, accumulators: this.specs.map(newAccumulator) }; this.groups.set(id, group); }
    } else group = this.totals;
    group.rows++;
    this.specs.forEach((spec, position) => { if (spec.column) feed(group.accumulators[position]!, spec, row[spec.column.index] as Cell, kinds[spec.column.index]!); else group.accumulators[position]!.count++; });
  }

  private list(row: Cell[], kinds: readonly Kind[], sheetRow: number): void {
    const by = this.listOrder;
    const entry = { row: sheetRow, values: this.selected.map((column) => (kinds[column.index] === "empty" ? null : row[column.index]!)), sort: by ? row[by.index]! : null };
    if (by) { if (this.listed.length < ROW_CAP) this.listed.push(entry); }
    else if (this.listed.length <= this.input.limit) this.listed.push(entry);
  }

  result(): Record<string, unknown> {
    const base = { rowsScanned: this.scanned };
    if (this.input.mode === "profile") return { ...base, columns: this.profiled.map((column) => this.describe(column)), notes: this.notes };
    const notes = [...this.notes];
    if (this.errorCells) notes.push(`${this.errorCells} cells in the matching rows show an Excel error; they are left out of the aggregates.`);
    if (this.grouping) return { ...base, matched: this.matched, ...this.groupResult(notes), notes };
    if (this.specs.length) return { ...base, matched: this.matched, totals: this.valuesOf(this.totals, notes), notes };
    const by = this.input.orderBy, direction = by?.direction === "asc" ? 1 : -1;
    const rows = by ? [...this.listed].sort((a, b) => direction * compareCells(a.sort, b.sort)) : this.listed;
    const shown = rows.slice(0, this.input.limit);
    return { ...base, matched: this.matched, columns: this.selected.map((column) => column.name), rowNumbers: shown.map((entry) => entry.row), rows: shown.map((entry) => entry.values), hasMore: this.matched > shown.length, notes };
  }

  private valuesOf(group: Group, notes: string[]): Record<string, Cell> {
    const out: Record<string, Cell> = {};
    this.specs.forEach((spec, position) => {
      const accumulator = group.accumulators[position]!;
      out[spec.name] = resultOf(accumulator, spec);
      if (accumulator.ignored && ["sum", "avg", "min", "max"].includes(spec.fn) && !group.key.length) notes.push(`${spec.name}: ${accumulator.ignored} non-numeric cells were ignored.`);
      if (accumulator.capped && !group.key.length) notes.push(`${spec.name}: more than ${DISTINCT_CAP} distinct values; the count stops there.`);
    });
    return out;
  }

  private groupResult(notes: string[]): Record<string, unknown> {
    const rows = this.input.orderBy;
    const names = [...this.groupColumns.map((column) => column.name), "rows", ...this.specs.map((spec) => spec.name)];
    const sortName = rows ? names.find((name) => name.toLowerCase() === rows.by.trim().toLowerCase()) ?? this.resolveSortName(rows.by) : this.specs[0]?.name ?? "rows";
    const direction = rows?.direction === "asc" ? 1 : -1;
    const entries = [...this.groups.values()].map((group) => {
      const record: Record<string, Cell> = {};
      this.groupColumns.forEach((column, position) => { record[column.name] = group.key[position] ?? null; });
      record.rows = group.rows;
      Object.assign(record, this.valuesOf(group, []));
      return record;
    });
    if (!names.includes(sortName)) throw new AnalyzeError(`Cannot order by '${rows?.by}'. Use one of: ${names.join(", ")}.`);
    entries.sort((a, b) => direction * compareCells(a[sortName] ?? null, b[sortName] ?? null));
    if (this.groupsCapped) notes.push(`More than ${GROUP_CAP} groups: the groups after the first ${GROUP_CAP} seen are not counted.`);
    return { groupCount: entries.length, groups: entries.slice(0, this.input.limit), hasMore: entries.length > this.input.limit };
  }

  private resolveSortName(reference: string): string {
    try { return resolveColumn(this.columns, reference).name; } catch { return reference; }
  }

  private describe(column: AnalyzeColumn): Record<string, unknown> {
    const profile = this.profiles.get(column.index)!;
    const total = profile.nonEmpty;
    const share = (count: number) => (total ? count / total : 0);
    const type = !total ? "empty" : share(profile.numbers) >= 0.95 ? "number" : share(profile.texts) >= 0.95 ? "text" : share(profile.booleans) >= 0.95 ? "boolean" : "mixed";
    const out: Record<string, unknown> = { name: column.name, letter: column.letter, type, filled: total, blank: profile.blank };
    if (profile.errors) out.errors = profile.errors;
    if (type === "mixed") out.mix = { numbers: profile.numbers, text: profile.texts, booleans: profile.booleans };
    if (column.dateLike) out.dateLike = true;
    if (profile.numbers) Object.assign(out, { min: profile.min, max: profile.max, mean: clean(profile.sum / profile.numbers) });
    out.distinct = profile.capped ? `more than ${DISTINCT_CAP}` : profile.values.size;
    if (!profile.capped && profile.values.size > 0 && profile.values.size <= (type === "number" ? 12 : DISTINCT_CAP)) {
      out.top = [...profile.values.values()].sort((a, b) => b.count - a.count).slice(0, 5).map((entry) => ({ value: entry.value, count: entry.count }));
    }
    out.sample = profile.sample;
    return out;
  }
}

function compareCells(a: Cell, b: Cell): number {
  if (a === null || a === undefined) return b === null || b === undefined ? 0 : -1;
  if (b === null || b === undefined) return 1;
  if (typeof a === "number" && typeof b === "number") return a - b;
  return String(a).toLowerCase().localeCompare(String(b).toLowerCase());
}
