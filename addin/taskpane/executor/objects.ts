import {
  CHART_TYPES,
  type GetObjectsInput,
  type ModifyChartInput,
  type ModifyPivotTableInput,
  type ModifyTableInput,
  type SheetRef,
} from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, stripSheet } from "./common.js";

const pascal = (text: string) => text.charAt(0).toUpperCase() + text.slice(1);
const suffix = () => Math.random().toString(36).slice(2, 7);

/** `Sheet1!A1:D10` / `'My Sheet'!A1:D10` -> that range; a bare address lives on `fallback`. */
export function rangeFromSource(context: Excel.RequestContext, fallback: Excel.Worksheet, source: string): Excel.Range {
  const bang = source.lastIndexOf("!");
  if (bang === -1) return fallback.getRange(source);
  const rawName = source.slice(0, bang);
  const name = rawName.startsWith("'") && rawName.endsWith("'") ? rawName.slice(1, -1).replace(/''/g, "'") : rawName;
  return context.workbook.worksheets.getItem(name).getRange(source.slice(bang + 1));
}

/** The sheet to act on when the caller gave none, searching every sheet for `find` (charts are per-sheet). */
async function locateChart(context: Excel.RequestContext, ref: SheetRef | undefined, name: string): Promise<{ sheet: Excel.Worksheet; chart: Excel.Chart }> {
  const candidates: Excel.Worksheet[] = [];
  if (ref) candidates.push(resolveSheet(context, ref));
  else {
    context.workbook.worksheets.load("items/name");
    await context.sync();
    candidates.push(...context.workbook.worksheets.items);
  }
  const probes = candidates.map((sheet) => {
    const chart = sheet.charts.getItemOrNullObject(name);
    chart.load("isNullObject");
    return { sheet, chart };
  });
  await context.sync();
  const hit = probes.find((p) => !p.chart.isNullObject);
  if (!hit) {
    throw new ToolFailure("invalid_arguments", `chart '${name}' not found${ref ? " on that sheet" : ""}; list charts with get_objects`);
  }
  return hit;
}

/* -------------------------------------------------------------------------- */
/* get_objects                                                                 */
/* -------------------------------------------------------------------------- */

export async function getObjects(input: GetObjectsInput): Promise<unknown> {
  return Excel.run(async (context) => {
    let sheets: Excel.Worksheet[];
    if (input.sheet) sheets = [resolveSheet(context, input.sheet)];
    else {
      context.workbook.worksheets.load("items/name,items/visibility");
      await context.sync();
      sheets = context.workbook.worksheets.items.filter((s) => s.visibility === "Visible");
    }
    sheets.forEach((sheet) => {
      sheet.load("name");
      sheet.charts.load("items/name,items/chartType,items/left,items/top,items/width,items/height");
      sheet.pivotTables.load("items/name");
      sheet.tables.load("items/name,items/style,items/showTotals,items/showHeaders,items/showFilterButton,items/showBandedRows");
    });
    await context.sync();

    const keep = (name: string) => !input.name || name === input.name;
    const charts = sheets.flatMap((sheet) => sheet.charts.items.filter((c) => keep(c.name)).map((chart) => ({ sheet, chart })));
    const pivots = sheets.flatMap((sheet) => sheet.pivotTables.items.filter((p) => keep(p.name)).map((pt) => ({ sheet, pt })));
    const tables = sheets.flatMap((sheet) => sheet.tables.items.filter((t) => keep(t.name)).map((table) => ({ sheet, table })));

    charts.forEach(({ chart }) => chart.title.load("text,visible"));
    const tableRanges = tables.map(({ table }) => {
      const range = table.getRange();
      range.load("address");
      return range;
    });
    pivots.forEach(({ pt }) => {
      pt.rowHierarchies.load("items/name");
      pt.columnHierarchies.load("items/name");
      pt.filterHierarchies.load("items/name");
      pt.dataHierarchies.load("items/name,items/summarizeBy");
    });
    const pivotRanges = pivots.map(({ pt }) => {
      const range = pt.layout.getRange();
      range.load("address");
      return range;
    });
    await context.sync();

    return {
      charts: charts.map(({ sheet, chart }) => ({
        sheet: sheet.name,
        name: chart.name,
        chartType: chart.chartType,
        title: chart.title.visible ? chart.title.text : null,
        left: Math.round(chart.left),
        top: Math.round(chart.top),
        width: Math.round(chart.width),
        height: Math.round(chart.height),
      })),
      pivotTables: pivots.map(({ sheet, pt }, i) => ({
        sheet: sheet.name,
        name: pt.name,
        range: stripSheet(pivotRanges[i]!.address),
        rows: pt.rowHierarchies.items.map((h) => h.name),
        columns: pt.columnHierarchies.items.map((h) => h.name),
        filters: pt.filterHierarchies.items.map((h) => h.name),
        values: pt.dataHierarchies.items.map((h) => ({ field: h.name, summarizeBy: h.summarizeBy })),
      })),
      tables: tables.map(({ sheet, table }, i) => ({
        sheet: sheet.name,
        name: table.name,
        range: stripSheet(tableRanges[i]!.address),
        style: table.style,
        showTotals: table.showTotals,
        showHeaders: table.showHeaders,
        showFilterButton: table.showFilterButton,
        bandedRows: table.showBandedRows,
      })),
    };
  });
}

/* -------------------------------------------------------------------------- */
/* Charts                                                                      */
/* -------------------------------------------------------------------------- */

function styleChart(chart: Excel.Chart, input: ModifyChartInput): void {
  if (input.title !== undefined) {
    chart.title.visible = input.title !== "";
    if (input.title !== "") chart.title.text = input.title;
  }
  if (input.anchor) chart.setPosition(input.anchor);
  if (input.width !== undefined) chart.width = input.width;
  if (input.height !== undefined) chart.height = input.height;
  if (input.legend) {
    chart.legend.visible = input.legend !== "none";
    if (input.legend !== "none") chart.legend.position = pascal(input.legend) as Excel.ChartLegendPosition;
  }
  if (input.xAxisTitle !== undefined) {
    chart.axes.categoryAxis.title.visible = true;
    chart.axes.categoryAxis.title.text = input.xAxisTitle;
  }
  if (input.yAxisTitle !== undefined) {
    chart.axes.valueAxis.title.visible = true;
    chart.axes.valueAxis.title.text = input.yAxisTitle;
  }
}

export async function modifyChart(input: ModifyChartInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const seriesBy = { auto: Excel.ChartSeriesBy.auto, rows: Excel.ChartSeriesBy.rows, columns: Excel.ChartSeriesBy.columns }[input.seriesBy];

    if (input.operation === "create") {
      const sheet = resolveSheet(context, input.sheet);
      sheet.load("name");
      const name = input.name ?? `Chart_${suffix()}`;
      const clash = sheet.charts.getItemOrNullObject(name);
      clash.load("isNullObject");
      await context.sync();
      if (!clash.isNullObject) throw new ToolFailure("invalid_arguments", `a chart named '${name}' already exists on this sheet`);

      const chart = sheet.charts.add(CHART_TYPES[input.chartType!] as Excel.ChartType, rangeFromSource(context, sheet, input.source!), seriesBy);
      chart.name = name;
      styleChart(chart, input);
      await context.sync();
      return { operation: "create", sheet: sheet.name, name, chartType: input.chartType };
    }

    const { sheet, chart } = await locateChart(context, input.sheet, input.name!);
    sheet.load("name");
    if (input.operation === "delete") {
      chart.delete();
      await context.sync();
      return { operation: "delete", sheet: sheet.name, name: input.name };
    }

    if (input.chartType) chart.chartType = CHART_TYPES[input.chartType] as Excel.ChartType;
    if (input.source) chart.setData(rangeFromSource(context, sheet, input.source), seriesBy);
    styleChart(chart, input);
    if (input.newName) chart.name = input.newName;
    await context.sync();
    return { operation: "update", sheet: sheet.name, name: input.newName ?? input.name };
  });
}

/* -------------------------------------------------------------------------- */
/* Pivot tables                                                                */
/* -------------------------------------------------------------------------- */

function assertFields(requested: string[], available: string[]): void {
  const missing = requested.filter((f) => !available.includes(f));
  if (missing.length > 0) {
    throw new ToolFailure("invalid_arguments", `unknown pivot field(s): ${missing.join(", ")}`, { availableFields: available });
  }
}

export async function modifyPivotTable(input: ModifyPivotTableInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const workbookPivots = context.workbook.pivotTables;
    const existing = workbookPivots.getItemOrNullObject(input.name);
    existing.load("isNullObject");
    await context.sync();

    if (input.operation === "create") {
      if (!existing.isNullObject) throw new ToolFailure("invalid_arguments", `a pivot table named '${input.name}' already exists`);
      const sheet = resolveSheet(context, input.sheet);
      sheet.load("name");
      const source = rangeFromSource(context, sheet, input.source!);
      const header = source.getRow(0);
      header.load("values");
      await context.sync();
      const fields = (header.values[0] ?? []).map((v) => String(v));
      assertFields([...(input.rows ?? []), ...(input.columns ?? []), ...(input.filters ?? []), ...input.values!.map((v) => v.field)], fields);

      const pt = sheet.pivotTables.add(input.name, source, sheet.getRange(input.destination!));
      addFields(pt, input);
      const layout = pt.layout.getRange();
      layout.load("address");
      await context.sync();
      return { operation: "create", sheet: sheet.name, name: input.name, range: stripSheet(layout.address) };
    }

    if (existing.isNullObject) throw new ToolFailure("invalid_arguments", `pivot table '${input.name}' not found; list them with get_objects`);
    if (input.operation === "delete") {
      existing.delete();
      await context.sync();
      return { operation: "delete", name: input.name };
    }

    const groups = [
      ["rows", existing.rowHierarchies] as const,
      ["columns", existing.columnHierarchies] as const,
      ["filters", existing.filterHierarchies] as const,
      ["values", existing.dataHierarchies] as const,
    ].filter(([key]) => input[key] !== undefined);

    existing.hierarchies.load("items/name");
    groups.forEach(([, collection]) => collection.load("items"));
    await context.sync();
    assertFields(
      [...(input.rows ?? []), ...(input.columns ?? []), ...(input.filters ?? []), ...(input.values ?? []).map((v) => v.field)],
      existing.hierarchies.items.map((h) => h.name),
    );
    for (const [key, collection] of groups) {
      for (const item of collection.items) {
        if (key === "values") existing.dataHierarchies.remove(item as Excel.DataPivotHierarchy);
        else (collection as Excel.RowColumnPivotHierarchyCollection).remove(item as Excel.RowColumnPivotHierarchy);
      }
    }
    await context.sync();
    addFields(existing, input);
    if (input.newName) existing.name = input.newName;
    await context.sync();
    return { operation: "update", name: input.newName ?? input.name };
  });
}

function addFields(pt: Excel.PivotTable, input: ModifyPivotTableInput): void {
  for (const field of input.rows ?? []) pt.rowHierarchies.add(pt.hierarchies.getItem(field));
  for (const field of input.columns ?? []) pt.columnHierarchies.add(pt.hierarchies.getItem(field));
  for (const field of input.filters ?? []) pt.filterHierarchies.add(pt.hierarchies.getItem(field));
  for (const value of input.values ?? []) {
    const hierarchy = pt.dataHierarchies.add(pt.hierarchies.getItem(value.field));
    hierarchy.summarizeBy = pascal(value.summarizeBy) as Excel.AggregationFunction;
    if (value.label) hierarchy.name = value.label;
  }
}

/* -------------------------------------------------------------------------- */
/* Excel Tables                                                                */
/* -------------------------------------------------------------------------- */

export async function modifyTable(input: ModifyTableInput): Promise<unknown> {
  return Excel.run(async (context) => {
    if (input.operation === "create") {
      const sheet = resolveSheet(context, input.sheet);
      sheet.load("name");
      if (input.name) {
        const clash = context.workbook.tables.getItemOrNullObject(input.name);
        clash.load("isNullObject");
        await context.sync();
        if (!clash.isNullObject) throw new ToolFailure("invalid_arguments", `a table named '${input.name}' already exists`);
      }
      const table = sheet.tables.add(sheet.getRange(input.range!), input.hasHeaders);
      if (input.name) table.name = input.name;
      styleTable(table, input);
      table.load("name");
      const range = table.getRange();
      range.load("address");
      await context.sync();
      return { operation: "create", sheet: sheet.name, name: table.name, range: stripSheet(range.address) };
    }

    const table = context.workbook.tables.getItemOrNullObject(input.name!);
    table.load("isNullObject,name");
    await context.sync();
    if (table.isNullObject) throw new ToolFailure("invalid_arguments", `table '${input.name}' not found; list them with get_objects`);

    if (input.operation === "delete") {
      if (input.deleteData) table.delete();
      else table.convertToRange();
      await context.sync();
      return { operation: "delete", name: input.name, dataKept: !input.deleteData };
    }

    if (input.resizeTo) table.resize(input.resizeTo);
    styleTable(table, input);
    if (input.newName) table.name = input.newName;
    await context.sync();
    return { operation: "update", name: input.newName ?? input.name };
  });
}

function styleTable(table: Excel.Table, input: ModifyTableInput): void {
  if (input.style) table.style = input.style;
  if (input.showTotals !== undefined) table.showTotals = input.showTotals;
  if (input.bandedRows !== undefined) table.showBandedRows = input.bandedRows;
  if (input.showFilterButton !== undefined) table.showFilterButton = input.showFilterButton;
}
