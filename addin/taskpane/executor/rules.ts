import type {
  CfFormat,
  CfRule,
  ConditionalFormatInput,
  DataValidationInput,
  RuleOperator,
  ValidationRule,
} from "../../../protocol/src/index.js";
import { ToolFailure, resolveSheet, stripSheet } from "./common.js";

const pascal = (text: string) => text.charAt(0).toUpperCase() + text.slice(1);

/* -------------------------------------------------------------------------- */
/* Pure builders (unit-tested): our schema -> Office.js rule objects            */
/* -------------------------------------------------------------------------- */

/** Formula text for a rule bound: numbers and `=...` pass through, plain text is quoted. */
export function ruleFormula(value: string | number): string {
  if (typeof value === "number") return `=${value}`;
  if (value.startsWith("=")) return value;
  return /^-?\d+(\.\d+)?$/.test(value) ? `=${value}` : `="${value.replace(/"/g, '""')}"`;
}

export function buildCellValueRule(rule: Extract<CfRule, { type: "cellValue" }>): Excel.ConditionalCellValueRule {
  return {
    formula1: ruleFormula(rule.value),
    ...(rule.value2 !== undefined ? { formula2: ruleFormula(rule.value2) } : {}),
    operator: pascal(rule.operator) as Excel.ConditionalCellValueOperator,
  };
}

const TEXT_OPERATOR = { contains: "Contains", notContains: "NotContains", beginsWith: "BeginsWith", endsWith: "EndsWith" } as const;
const TOP_BOTTOM = { topItems: "TopItems", topPercent: "TopPercent", bottomItems: "BottomItems", bottomPercent: "BottomPercent" } as const;

function applyFormat(target: { fill: { color: string }; font: { color: string; bold: boolean; italic: boolean; strikethrough: boolean } }, format: CfFormat): void {
  if (format.fillColor) target.fill.color = format.fillColor;
  if (format.fontColor) target.font.color = format.fontColor;
  if (format.bold !== undefined) target.font.bold = format.bold;
  if (format.italic !== undefined) target.font.italic = format.italic;
  if (format.strikethrough !== undefined) target.font.strikethrough = format.strikethrough;
}

function addRule(range: Excel.Range, rule: CfRule): void {
  const types = Excel.ConditionalFormatType;
  switch (rule.type) {
    case "cellValue": {
      const cf = range.conditionalFormats.add(types.cellValue);
      cf.cellValue.rule = buildCellValueRule(rule);
      applyFormat(cf.cellValue.format as never, rule.format);
      return;
    }
    case "containsText": {
      const cf = range.conditionalFormats.add(types.containsText);
      cf.textComparison.rule = { operator: TEXT_OPERATOR[rule.match] as Excel.ConditionalTextOperator, text: rule.text };
      applyFormat(cf.textComparison.format as never, rule.format);
      return;
    }
    case "custom": {
      const cf = range.conditionalFormats.add(types.custom);
      cf.custom.rule.formula = rule.formula.startsWith("=") ? rule.formula : `=${rule.formula}`;
      applyFormat(cf.custom.format as never, rule.format);
      return;
    }
    case "colorScale": {
      const cf = range.conditionalFormats.add(types.colorScale);
      cf.colorScale.criteria = {
        minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: rule.minColor },
        ...(rule.midColor ? { midpoint: { type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50", color: rule.midColor } } : {}),
        maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: rule.maxColor },
      };
      return;
    }
    case "dataBar": {
      const cf = range.conditionalFormats.add(types.dataBar);
      if (rule.color) cf.dataBar.positiveFormat.fillColor = rule.color;
      if (rule.negativeColor) cf.dataBar.negativeFormat.fillColor = rule.negativeColor;
      cf.dataBar.showDataBarOnly = !rule.showValue;
      return;
    }
    case "iconSet": {
      const cf = range.conditionalFormats.add(types.iconSet);
      cf.iconSet.style = pascal(rule.style) as Excel.IconSet;
      cf.iconSet.reverseIconOrder = rule.reverse;
      cf.iconSet.showIconOnly = !rule.showValue;
      return;
    }
    case "topBottom": {
      const cf = range.conditionalFormats.add(types.topBottom);
      cf.topBottom.rule = { rank: rule.rank, type: TOP_BOTTOM[rule.kind] as Excel.ConditionalTopBottomCriterionType };
      applyFormat(cf.topBottom.format as never, rule.format);
      return;
    }
    case "duplicates": {
      const cf = range.conditionalFormats.add(types.presetCriteria);
      cf.preset.rule = { criterion: (rule.unique ? "UniqueValues" : "DuplicateValues") as Excel.ConditionalFormatPresetCriterion };
      applyFormat(cf.preset.format as never, rule.format);
      return;
    }
  }
}

export async function conditionalFormat(input: ConditionalFormatInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = input.address ? sheet.getRange(input.address) : sheet.getRange();

    if (input.operation === "add") {
      addRule(range, input.rule!);
      range.load("address");
      await context.sync();
      return { sheet: sheet.name, address: stripSheet(range.address), added: input.rule!.type };
    }
    if (input.operation === "clear") {
      range.conditionalFormats.clearAll();
      range.load("address");
      await context.sync();
      return { sheet: sheet.name, address: stripSheet(range.address), cleared: true };
    }

    range.conditionalFormats.load("items/type,items/priority,items/stopIfTrue");
    await context.sync();
    const where = range.conditionalFormats.items.map((cf) => {
      const r = cf.getRange();
      r.load("address");
      return r;
    });
    if (where.length > 0) await context.sync();
    return {
      sheet: sheet.name,
      rules: range.conditionalFormats.items.map((cf, i) => ({ type: cf.type, priority: cf.priority, address: stripSheet(where[i]!.address) })),
    };
  });
}

/* -------------------------------------------------------------------------- */
/* Data validation                                                             */
/* -------------------------------------------------------------------------- */

export function buildValidationRule(rule: ValidationRule): Excel.DataValidationRule {
  if (rule.type === "list") {
    if (Array.isArray(rule.source)) {
      if (rule.source.some((item) => item.includes(","))) {
        throw new ToolFailure("invalid_arguments", "list items cannot contain commas; use a range source instead");
      }
      return { list: { inCellDropDown: rule.inCellDropDown, source: rule.source.join(",") } };
    }
    return { list: { inCellDropDown: rule.inCellDropDown, source: rule.source } };
  }
  if (rule.type === "custom") return { custom: { formula: rule.formula } };

  const bounds = {
    formula1: rule.value1,
    ...(rule.value2 !== undefined ? { formula2: rule.value2 } : {}),
    operator: pascal(rule.operator),
  };
  const key = { whole: "wholeNumber", decimal: "decimal", date: "date", time: "time", textLength: "textLength" }[rule.type];
  return { [key]: bounds } as Excel.DataValidationRule;
}

export async function dataValidation(input: DataValidationInput): Promise<unknown> {
  return Excel.run(async (context) => {
    const sheet = resolveSheet(context, input.sheet);
    sheet.load("name");
    const range = sheet.getRange(input.address);
    const validation = range.dataValidation;
    range.load("address");

    if (input.operation === "clear") {
      validation.clear();
      await context.sync();
      return { sheet: sheet.name, address: stripSheet(range.address), cleared: true };
    }
    if (input.operation === "get") {
      validation.load("type,rule,ignoreBlanks,prompt,errorAlert,valid");
      await context.sync();
      return {
        sheet: sheet.name,
        address: stripSheet(range.address),
        type: validation.type,
        rule: validation.rule,
        ignoreBlanks: validation.ignoreBlanks,
        prompt: validation.prompt,
        errorAlert: validation.errorAlert,
      };
    }

    validation.rule = buildValidationRule(input.rule!);
    validation.ignoreBlanks = input.ignoreBlanks;
    if (input.prompt) validation.prompt = { showPrompt: true, title: input.prompt.title ?? "", message: input.prompt.message };
    if (input.errorAlert) {
      validation.errorAlert = {
        showAlert: true,
        style: pascal(input.errorAlert.style) as Excel.DataValidationAlertStyle,
        title: input.errorAlert.title ?? "",
        message: input.errorAlert.message,
      };
    }
    await context.sync();
    return { sheet: sheet.name, address: stripSheet(range.address), type: input.rule!.type };
  });
}

export type { RuleOperator };
