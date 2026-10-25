import {
  columnToNumber,
  type FilterCriteria,
  type FilterOperator,
  type FilterRangeInput,
  type SortRangeInput,
} from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, stripSheet } from "./common.js";
import { loadGeometry } from "./write.js";

/** Relative 0-based column index of a sheet column letter inside a range, or a clear error. */
export function relativeColumn(letter: string, firstColumnIndex: number, columnCount: number, address: string): number {
  const index = columnToNumber(letter) - 1 - firstColumnIndex;
  if (index < 0 || index >= columnCount) {
    throw new ToolFailure("invalid_arguments", `column ${letter.toUpperCase()} is outside the range ${address}`);
  }
  return index;
}

export async function sortRange(input: SortRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    const geometry = await loadGeometry(context, range);
    const fields = input.keys.map((key) => ({
      key: relativeColumn(key.column, geometry.columnIndex, geometry.columnCount, input.address),
      ascending: key.ascending,
    }));
    range.sort.apply(fields, input.matchCase, input.hasHeaders, Excel.SortOrientation.rows);
    range.load("address");
    await context.sync();
    return { sheet: sheet.name, address: stripSheet(range.address), sortedBy: input.keys.map((k) => `${k.column.toUpperCase()} ${k.ascending ? "asc" : "desc"}`) };
  });
}

/* -------------------------------------------------------------------------- */
/* Filters                                                                     */
/* -------------------------------------------------------------------------- */

/** Excel custom-filter criterion text for one condition. */
export function filterCriterion(operator: FilterOperator, value?: string | number): string {
  const v = value === undefined ? "" : String(value);
  switch (operator) {
    case "equals":
      return `=${v}`;
    case "notEquals":
      return `<>${v}`;
    case "greaterThan":
      return `>${v}`;
    case "greaterThanOrEqual":
      return `>=${v}`;
    case "lessThan":
      return `<${v}`;
    case "lessThanOrEqual":
      return `<=${v}`;
    case "contains":
      return `=*${v}*`;
    case "notContains":
      return `<>*${v}*`;
    case "beginsWith":
      return `=${v}*`;
    case "endsWith":
      return `=*${v}`;
    case "blank":
      return "=";
    case "nonBlank":
      return "<>";
  }
}

export function buildFilterCriteria(criteria: FilterCriteria): Excel.FilterCriteria {
  if ("values" in criteria) return { filterOn: Excel.FilterOn.values, values: criteria.values };
  if ("top" in criteria) {
    return { filterOn: pascalOn(criteria.top.kind) as Excel.FilterOn, criterion1: String(criteria.top.n) };
  }
  const [first, second] = criteria.conditions;
  for (const c of criteria.conditions) {
    if (!["blank", "nonBlank"].includes(c.operator) && c.value === undefined) {
      throw new ToolFailure("invalid_arguments", `filter condition '${c.operator}' needs a value`);
    }
  }
  return {
    filterOn: Excel.FilterOn.custom,
    criterion1: filterCriterion(first!.operator, first!.value),
    ...(second ? { criterion2: filterCriterion(second.operator, second.value), operator: criteria.join === "or" ? Excel.FilterOperator.or : Excel.FilterOperator.and } : {}),
  };
}

const pascalOn = (kind: string) => kind.charAt(0).toUpperCase() + kind.slice(1);

/** Filtered range and how many data rows are visible; best effort, never fails the call. */
async function describeFilter(context: Excel.RequestContext, filter: Excel.AutoFilter): Promise<Record<string, unknown>> {
  try {
    const range = filter.getRangeOrNullObject();
    range.load("isNullObject,address,rowCount");
    await context.sync();
    if (range.isNullObject) return { enabled: false };
    const out: Record<string, unknown> = { enabled: true, filteredRange: stripSheet(range.address), dataRows: Math.max(0, range.rowCount - 1) };
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
      const view = range.getVisibleView();
      view.load("rowCount");
      await context.sync();
      out.visibleRows = Math.max(0, view.rowCount - 1);
    }
    return out;
  } catch {
    return {};
  }
}

export async function filterRange(input: FilterRangeInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const filter = sheet.autoFilter;

    switch (input.operation) {
      case "clear":
        filter.clearCriteria();
        break;
      case "reapply":
        filter.reapply();
        break;
      case "remove":
        filter.remove();
        break;
      case "apply": {
        const range = sheet.getRange(input.address!);
        const geometry = await loadGeometry(context, range);
        // Re-applying to a different range needs the old filter removed; the same range keeps other columns' criteria.
        filter.load("enabled");
        const current = filter.getRangeOrNullObject();
        current.load("isNullObject,address");
        range.load("address");
        await context.sync();
        if (filter.enabled && !current.isNullObject && stripSheet(current.address) !== stripSheet(range.address)) {
          filter.remove();
        }
        if (input.criteria && input.column) {
          const column = relativeColumn(input.column, geometry.columnIndex, geometry.columnCount, input.address!);
          filter.apply(range, column, buildFilterCriteria(input.criteria));
        } else {
          filter.apply(range);
        }
        break;
      }
    }
    await context.sync();
    return {
      sheet: sheet.name,
      operation: input.operation,
      ...(input.address ? { address: input.address } : {}),
      ...(input.column && input.criteria ? { column: input.column, criteria: input.criteria } : {}),
      ...(input.operation === "remove" ? { enabled: false } : await describeFilter(context, filter)),
    };
  });
}
