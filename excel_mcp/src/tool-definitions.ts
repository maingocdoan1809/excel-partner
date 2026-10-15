import { EIGHT_SITE_SHAPES, FOUR_SITE_SHAPES } from "../../protocol/src/index.js";

export interface JsonSchema {
  type?: string | string[];
  description?: string;
  properties?: Record<string, JsonSchema>;
  items?: JsonSchema;
  required?: string[];
  enum?: string[];
  default?: unknown;
  minimum?: number;
  maximum?: number;
  additionalProperties?: boolean | JsonSchema;
  minItems?: number;
  pattern?: string;
}

export interface ToolDefinition {
  name: string;
  title: string;
  description: string;
  inputSchema: JsonSchema;
  annotations: {
    readOnlyHint: boolean;
    destructiveHint: boolean;
    idempotentHint: boolean;
  };
}

const sheetRef: JsonSchema = {
  type: "object",
  description: "Target sheet. Prefer `name`; `id` survives renames.",
  properties: {
    name: { type: "string", description: "Sheet display name." },
    id: { type: "string", description: "Office.js sheet id." },
  },
  additionalProperties: false,
};

const cellValue: JsonSchema = {
  type: ["string", "number", "boolean", "null"],
  description: "A single cell value.",
};

const color: JsonSchema = { type: "string", pattern: "^#[0-9A-Fa-f]{6}$" };

/**
 * JSON Schemas advertised to MCP clients. Kept hand-written (draft-07 subset)
 * so the tool surface is explicit and independent of the SDK's zod bridge;
 * runtime validation still uses the shared zod schemas in `protocol`.
 */
const READ = { readOnlyHint: true, destructiveHint: false, idempotentHint: true } as const;
const WRITE = { readOnlyHint: false, destructiveHint: true, idempotentHint: false } as const;

const overwriteFlag: JsonSchema = {
  type: "boolean",
  default: false,
  description:
    "Writing over non-empty cells fails unless this is true. Leave it unset on the first attempt; " +
    "set it only after the user confirmed the overwrite or explicitly asked to replace data.",
};

const rangeAddress: JsonSchema = {
  type: "string",
  description: "A1 address, e.g. `A1` or `B2:D10`.",
};

/**
 * JSON Schemas advertised to MCP clients. Kept hand-written (draft-07 subset)
 * so the tool surface is explicit and independent of the SDK's zod bridge;
 * runtime validation still uses the shared zod schemas in `protocol`
 * (`tests/protocol.test.ts` checks the two stay in sync).
 */
const shapeTypes = [...FOUR_SITE_SHAPES, ...EIGHT_SITE_SHAPES, "TextBox"];

const shapeStyle: Record<string, JsonSchema> = {
  fillColor: color,
  lineColor: { type: "string", description: "Outline colour `#RRGGBB`, or `none` for no outline." },
  lineWeight: { type: "number", minimum: 0.25, maximum: 20, description: "Points. Excel only takes whole points; fractions are rounded (min 1)." },
  fontColor: color,
  fontSize: { type: "number", minimum: 1, maximum: 409 },
  fontName: { type: "string", description: "Font family, e.g. Arial." },
  bold: { type: "boolean" },
  horizontalAlignment: { type: "string", enum: ["Left", "Center", "Right"] },
};

const shapePlacement: Record<string, JsonSchema> = {
  anchor: {
    type: "string",
    description:
      "Cell range the shape covers, e.g. `M10:M11` (preferred: keeps shapes on the sheet grid). " +
      "Alternatively give all of left/top/width/height in points.",
  },
  left: { type: "number" },
  top: { type: "number" },
  width: { type: "number" },
  height: { type: "number" },
};

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "get_selection",
    title: "Get selection",
    description:
      "READ. The range the user currently has selected: sheet, address, values and column headers. " +
      "Call this first to understand what the user is pointing at.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: READ,
  },
  {
    name: "get_sheets",
    title: "List sheets",
    description: "READ. Every worksheet with visibility and used range.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: READ,
  },
  {
    name: "read_range",
    title: "Read range",
    description:
      "READ. Values (and optionally formulas) of an A1 range. Use format `csv` for large ranges: " +
      "it is much more compact than the default 2D array. Page big ranges with maxRows / offset.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        address: rangeAddress,
        includeFormulas: { type: "boolean", default: false },
        format: { type: "string", enum: ["grid", "csv"], default: "grid" },
        maxRows: { type: "integer", minimum: 1, maximum: 20000 },
        offset: { type: "integer", minimum: 0, default: 0 },
      },
      required: ["address"],
      additionalProperties: false,
    },
    annotations: READ,
  },
  {
    name: "search_data",
    title: "Search data",
    description:
      "READ. Find cells containing text across all visible sheets (or one sheet/range). Returns cell " +
      "addresses to follow up with read_range. Page with `offset` / `nextOffset`.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        address: { ...rangeAddress, description: "Optional A1 range limiting the search (needs `sheet`)." },
        searchTerm: { type: "string" },
        matchCase: { type: "boolean", default: false },
        wholeCell: { type: "boolean", default: false },
        limit: { type: "integer", minimum: 1, maximum: 500, default: 50 },
        offset: { type: "integer", minimum: 0, default: 0 },
      },
      required: ["searchTerm"],
      additionalProperties: false,
    },
    annotations: READ,
  },
  {
    name: "write_range",
    title: "Write values",
    description:
      "WRITE. Write a 2D array of values (strings starting with `=` become formulas). `address` must be a " +
      "single cell or exactly match the size of `values`; with mode `top_left` it is only the anchor. " +
      "Refuses to overwrite non-empty cells unless allowOverwrite is true. `notes` attaches cell notes " +
      "(e.g. a data source) in the same call.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        address: rangeAddress,
        values: {
          type: "array",
          minItems: 1,
          items: { type: "array", minItems: 1, items: cellValue },
        },
        mode: { type: "string", enum: ["values", "top_left"], default: "values" },
        allowOverwrite: overwriteFlag,
        notes: {
          type: "object",
          description: "Map of cell address to note text, e.g. {\"B2\": \"Source: ...\"}.",
          additionalProperties: { type: "string" },
        },
      },
      required: ["address", "values"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "set_formula",
    title: "Set formula",
    description:
      "WRITE. Set a formula in A1 notation. It is written to the top-left cell of `address` and filled " +
      "across the rest of the range, adjusting relative references (e.g. address `D2:D100`, formula " +
      "`B2*C2`). Refuses to overwrite non-empty cells unless allowOverwrite is true.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        address: rangeAddress,
        formula: { type: "string", description: "Formula with or without a leading `=`, e.g. `SUM(A1:A3)`." },
        language: { type: "string", enum: ["A1"], default: "A1" },
        allowOverwrite: overwriteFlag,
      },
      required: ["address", "formula"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "format_range",
    title: "Format range",
    description:
      "WRITE. Apply formatting (font, fill, number format, alignment, borders, column width / row height, " +
      "autofit) to a range. Only the listed properties change.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        address: rangeAddress,
        format: {
          type: "object",
          properties: {
            bold: { type: "boolean" },
            italic: { type: "boolean" },
            underline: { type: "boolean" },
            fontSize: { type: "number", minimum: 1, maximum: 409 },
            fontColor: color,
            fillColor: color,
            numberFormat: { type: "string", description: "e.g. `#,##0.00`, `0%`, `yyyy-mm-dd`." },
            horizontalAlignment: { type: "string", enum: ["Left", "Center", "Right", "Justify", "General"] },
            verticalAlignment: { type: "string", enum: ["Top", "Center", "Bottom"] },
            wrapText: { type: "boolean" },
            columnWidth: { type: "number" },
            rowHeight: { type: "number" },
            borders: { type: "string", enum: ["all", "outline", "none"] },
            autofitColumns: { type: "boolean" },
          },
          additionalProperties: false,
        },
      },
      required: ["address", "format"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "clear_range",
    title: "Clear range",
    description:
      "WRITE. Clear cells. Default `contents` removes values/formulas and keeps formatting; `formats` " +
      "removes only formatting; `all` removes both. Accepts `A1:C10`, whole columns `D:D` and rows `3:3`. " +
      "Only use when the user asked to delete or clear data.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        address: { type: "string", description: "A1 range, `D:D` or `3:3`." },
        clearType: { type: "string", enum: ["contents", "formats", "all"], default: "contents" },
      },
      required: ["address"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "insert_rows",
    title: "Insert rows",
    description: "WRITE. Insert blank rows above or below a row reference.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        at: { type: "string", description: "Row reference like `5` or `A5`." },
        count: { type: "integer", minimum: 1, maximum: 1000 },
        position: { type: "string", enum: ["above", "below"], default: "above" },
      },
      required: ["at", "count"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "modify_structure",
    title: "Modify sheet structure",
    description:
      "WRITE. insert/delete/hide/unhide/group/ungroup rows or columns starting at `reference` (row number like `5` or " +
      "column letters like `C`), freeze the first `count` rows/columns, unfreeze, or collapse/expand grouped " +
      "rows/columns with show_levels (count = outline level to show). Prefer GROUPING over hiding unless the user asks " +
      "to hide: groups show a +/- toggle. Deleting removes data; only do it when the user asked.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        operation: { type: "string", enum: ["insert", "delete", "hide", "unhide", "freeze", "unfreeze", "group", "ungroup", "show_levels"] },
        dimension: { type: "string", enum: ["rows", "columns"], description: "Required except for unfreeze." },
        reference: { type: "string", description: "Row number or column letters; not used by freeze/unfreeze." },
        count: { type: "integer", minimum: 1, maximum: 1000, default: 1 },
        position: { type: "string", enum: ["before", "after"], default: "before", description: "For insert." },
      },
      required: ["operation"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "list_shapes",
    title: "List shapes",
    description:
      "READ. Shapes, text boxes and connectors on a sheet: name, type, text and position in points. " +
      "Use it to find existing shapes before adding, connecting or moving.",
    inputSchema: { type: "object", properties: { sheet: sheetRef }, additionalProperties: false },
    annotations: READ,
  },
  {
    name: "add_shape",
    title: "Add shape",
    description:
      "WRITE. Draw a real Excel shape (rectangle, diamond/decision, terminator, input/output, ellipse, text " +
      "box ...) with optional text and style. Give it a unique `name` so add_connector can link it. Place " +
      "it over a cell range with `anchor` (e.g. flowchart column M, one or two rows per step).",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        name: { type: "string", description: "Unique shape name on the sheet, e.g. `step_1`." },
        type: {
          type: "string",
          enum: shapeTypes,
          description:
            "Flowchart: FlowChartTerminator (start/end), FlowChartProcess, FlowChartDecision, " +
            "FlowChartInputOutput, FlowChartPredefinedProcess, FlowChartDocument, FlowChartConnector. " +
            "General: Rectangle, RoundRectangle, Diamond, Ellipse, TextBox.",
        },
        text: { type: "string" },
        ...shapePlacement,
        ...shapeStyle,
      },
      required: ["type"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "add_connector",
    title: "Connect shapes",
    description:
      "WRITE. Draw a connector (arrow) that stays attached to two named shapes when they move. Sides " +
      "default to `auto` (the facing edges). Use `fromSide: right` etc. for decision branches. For label " +
      "text such as Yes/No, add a small TextBox next to the line.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        name: { type: "string" },
        from: { type: "string", description: "Name of the source shape." },
        to: { type: "string", description: "Name of the target shape." },
        fromSide: { type: "string", enum: ["auto", "top", "bottom", "left", "right"], default: "auto" },
        toSide: { type: "string", enum: ["auto", "top", "bottom", "left", "right"], default: "auto" },
        fromSite: { type: "integer", minimum: 0, description: "Raw connection-site index; overrides fromSide." },
        toSite: { type: "integer", minimum: 0, description: "Raw connection-site index; overrides toSide." },
        connectorType: { type: "string", enum: ["Straight", "Elbow", "Curve"], default: "Elbow" },
        arrow: { type: "string", enum: ["end", "both", "none", "diamond"], default: "end" },
        dash: { type: "string", enum: ["solid", "dash", "longDash", "dashDot", "dashDotDot", "dot"], default: "solid" },
        color,
        weight: { type: "number", minimum: 0.25, maximum: 20, description: "Points; rounded to whole points (min 1)." },
      },
      required: ["from", "to"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "update_shape",
    title: "Update shape",
    description:
      "WRITE. Change an existing shape by name: text, position/size, colours, font, or rename it. Only the " +
      "properties you pass change; connectors follow a moved shape.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        name: { type: "string" },
        newName: { type: "string" },
        text: { type: "string" },
        ...shapePlacement,
        ...shapeStyle,
      },
      required: ["name"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
  {
    name: "delete_shapes",
    title: "Delete shapes",
    description: "WRITE. Delete shapes or connectors by name. Only when the user asked to remove them.",
    inputSchema: {
      type: "object",
      properties: {
        sheet: sheetRef,
        names: { type: "array", minItems: 1, items: { type: "string" } },
      },
      required: ["names"],
      additionalProperties: false,
    },
    annotations: WRITE,
  },
];

import { ADVANCED_TOOL_DEFINITIONS } from "./tool-definitions-advanced.js";

/** Every tool definition: the hand-written core set plus the generated advanced set. */
export const ALL_TOOL_DEFINITIONS: ToolDefinition[] = [...TOOL_DEFINITIONS, ...ADVANCED_TOOL_DEFINITIONS];
