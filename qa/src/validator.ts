import fs from "node:fs";
import path from "node:path";
import { parse, type Node as AcornNode } from "acorn";
import { QaError, TC_ID_RE, isInside, sha256Text } from "./util.js";

/**
 * Static checks for AI-written Playwright scripts. This is a guardrail against mistakes and prompt-injected
 * code, not a sandbox: containment comes from the runner (scrubbed environment, host allowlist, approval hash).
 */
export interface Violation { rule: string; message: string; line?: number; column?: number }

export interface ValidationResult {
  ok: boolean;
  violations: Violation[];
  tcId?: string;
  reqId?: string;
  /** Project-relative (posix) paths of the page objects this file imports, transitively. */
  deps: string[];
}

export interface ValidateOptions {
  kind: "test" | "page";
  /** Absolute automation/ folder. */
  automationDir: string;
  /** Absolute path the source is (or will be) saved to. */
  file: string;
  allowedOrigins: string[];
  expectedTcId?: string;
}

export const MAX_SCRIPT_BYTES = 60_000;
const MAX_NODES = 25_000;
const ALLOWED_PACKAGE = "@playwright/test";
const ALLOWED_USE_KEYS = new Set(["role", "viewport", "locale", "timezoneId", "colorScheme"]);

const BANNED_IDENTIFIERS = new Map<string, string>([
  ["eval", "dynamic code"], ["Function", "dynamic code"], ["require", "CommonJS loading"], ["process", "Node process access"],
  ["globalThis", "global object access"], ["global", "global object access"], ["module", "module internals"], ["exports", "module internals"],
  ["fetch", "network calls outside the browser allowlist (use the odoo.rpc helper)"], ["WebSocket", "raw sockets"], ["XMLHttpRequest", "network calls outside the browser allowlist"],
  ["Worker", "workers"], ["SharedArrayBuffer", "shared memory"], ["Reflect", "reflection"], ["Proxy", "proxies"],
]);

const BANNED_MEMBERS = new Map<string, string>([
  ["route", "changing network routing would defeat the origin allowlist"], ["unroute", "changing network routing"], ["routeFromHAR", "changing network routing"],
  ["routeWebSocket", "changing network routing"], ["newContext", "contexts are created by the harness so the allowlist and session apply"],
  ["launch", "browsers are launched by the harness"], ["launchPersistentContext", "browsers are launched by the harness"], ["launchServer", "browsers are launched by the harness"],
  ["connect", "attaching to other browsers"], ["connectOverCDP", "attaching to other browsers"], ["newCDPSession", "raw DevTools protocol"], ["newBrowserCDPSession", "raw DevTools protocol"],
  ["cookies", "reading session cookies"], ["addCookies", "writing session cookies"], ["clearCookies", "tampering with session cookies"], ["storageState", "reading session tokens"],
  ["setExtraHTTPHeaders", "injecting credentials"], ["exposeFunction", "bridging page code to Node"], ["exposeBinding", "bridging page code to Node"],
  ["addInitScript", "page-wide injected scripts"], ["request", "the Node API client bypasses the browser allowlist (use the odoo.rpc helper)"],
  ["constructor", "constructor access can rebuild blocked globals"], ["__proto__", "prototype access"], ["prototype", "prototype access"],
  ["fetch", "network calls outside the browser allowlist"], ["_channel", "Playwright internals"], ["_connection", "Playwright internals"],
  ["_initializer", "Playwright internals"], ["_guid", "Playwright internals"], ["_wrapApiCall", "Playwright internals"], ["_platform", "Playwright internals"],
]);

const PASSWORD_RECEIVER = /pass(word|wd)?\b|\bpwd\b|secret|token|m[aậ]t\s*kh[aẩ]u/i;
const SECRET_LITERALS: Array<[RegExp, string]> = [
  [/\beyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\b/, "a JWT"],
  [/\bBearer\s+[A-Za-z0-9._~+/=-]{12,}/i, "a Bearer token"],
  [/-----BEGIN [A-Z ]*PRIVATE KEY-----/, "a private key"],
  [/\bAKIA[0-9A-Z]{16}\b/, "an AWS access key"],
  [/\b(?:ghp|gho|ghu|ghs|github_pat)_[A-Za-z0-9_]{20,}\b/, "a GitHub token"],
  [/^[A-Za-z0-9+/_=-]{40,}$/, "a long opaque string that looks like a key or token"],
];

type Ast = AcornNode & Record<string, any>;

function lineOf(node: Ast): { line: number; column: number } {
  const loc = node.loc?.start;
  return { line: loc?.line ?? 0, column: (loc?.column ?? 0) + 1 };
}

function isNode(value: unknown): value is Ast {
  return Boolean(value) && typeof value === "object" && typeof (value as { type?: unknown }).type === "string";
}

function walk(node: Ast, parent: Ast | null, visit: (node: Ast, parent: Ast | null) => void): void {
  visit(node, parent);
  for (const key of Object.keys(node)) {
    if (key === "loc" || key === "type") continue;
    const child = (node as Record<string, unknown>)[key];
    if (Array.isArray(child)) for (const item of child) { if (isNode(item)) walk(item, node, visit); }
    else if (isNode(child)) walk(child, node, visit);
  }
}

const stringValue = (node: Ast | undefined): string | null => {
  if (!node) return null;
  if (node.type === "Literal" && typeof node.value === "string") return node.value;
  if (node.type === "TemplateLiteral") return (node.quasis as Ast[]).map(quasi => quasi.value.cooked ?? "").join("\u0000");
  return null;
};

function memberName(node: Ast): string | null {
  if (!node.computed) return node.property?.type === "Identifier" ? node.property.name : node.property?.type === "PrivateIdentifier" ? `#${node.property.name}` : null;
  return node.property?.type === "Literal" && typeof node.property.value === "string" ? node.property.value : null;
}

/** Is this Identifier a reference to a variable (as opposed to a property name, label or import/export name)? */
function isReference(node: Ast, parent: Ast | null): boolean {
  if (!parent) return true;
  switch (parent.type) {
    case "MemberExpression": return parent.object === node || (parent.computed && parent.property === node);
    case "Property": return parent.computed ? parent.key === node || parent.value === node : parent.value === node;
    case "MethodDefinition": case "PropertyDefinition": return parent.computed && parent.key === node;
    case "LabeledStatement": case "BreakStatement": case "ContinueStatement": return false;
    case "ImportSpecifier": return false;
    case "ImportDefaultSpecifier": case "ImportNamespaceSpecifier": return false;
    case "ExportSpecifier": return parent.local === node;
    default: return true;
  }
}

function allowedComputed(property: Ast): boolean {
  if (property.type === "Literal") return typeof property.value === "number" || typeof property.value === "string";
  if (property.type === "Identifier") return true;
  if (property.type === "UnaryExpression") return property.operator === "-" && property.argument?.type === "Literal";
  if (property.type === "BinaryExpression") {
    const numeric = (n: Ast) => n.type === "Literal" ? typeof n.value === "number" : n.type === "Identifier";
    return ["-", "*", "%", "+"].includes(property.operator) && numeric(property.left) && numeric(property.right)
      && !(property.operator === "+" && property.left.type === "Identifier" && property.right.type === "Identifier");
  }
  return false;
}

/** Resolve an import written in `file` to an absolute path inside `_harness/` or `pages/`, or null. */
function resolveLocalImport(source: string, file: string, automationDir: string): { target: string; trusted: boolean } | null {
  if (!source.startsWith("./") && !source.startsWith("../")) return null;
  const target = path.resolve(path.dirname(file), source);
  const harness = path.join(automationDir, "_harness"), pages = path.join(automationDir, "pages");
  if (isInside(harness, target)) return { target, trusted: true };
  if (isInside(pages, target)) return { target, trusted: false };
  return null;
}

export function validateSource(source: string, options: ValidateOptions, seen = new Set<string>()): ValidationResult {
  const violations: Violation[] = [];
  const add = (rule: string, message: string, node?: Ast) => violations.push({ rule, message, ...(node ? lineOf(node) : {}) });
  const result: ValidationResult = { ok: false, violations, deps: [] };
  const text = source.replace(/^﻿/, "");
  const posixRelative = (absolute: string) => path.relative(path.dirname(options.automationDir), absolute).split(path.sep).join("/");

  if (Buffer.byteLength(text) > MAX_SCRIPT_BYTES) { add("size", `Script is larger than ${MAX_SCRIPT_BYTES / 1000} KB; split it into page objects.`); return result; }
  if (/\0/.test(text)) { add("encoding", "Script contains a NUL character."); return result; }

  let program: Ast;
  try {
    program = parse(text, { ecmaVersion: "latest", sourceType: "module", locations: true, allowHashBang: false }) as Ast;
  } catch (error) {
    const loc = (error as { loc?: { line: number; column: number } }).loc;
    violations.push({ rule: "syntax", message: `Not valid JavaScript: ${(error as Error).message.replace(/ \(\d+:\d+\)$/, "")}. Write plain JavaScript (ES modules), not TypeScript.`, line: loc?.line, column: loc ? loc.column + 1 : undefined });
    return result;
  }

  // Header and id checks.
  if (options.kind === "test") {
    const head = text.split(/\r?\n/).slice(0, 8).join("\n");
    const header = /^\/\/\s*TC:\s*([A-Za-z0-9_-]+)\s*\|\s*Req:\s*([^|\r\n]*?)\s*(?:\|[^\r\n]*)?$/m.exec(head);
    if (!header) add("header", "Start the file with: // TC: <id> | Req: <requirement id> (use Req: - when there is none).");
    else {
      result.tcId = header[1]; result.reqId = header[2] && header[2] !== "-" ? header[2] : undefined;
      if (!TC_ID_RE.test(result.tcId!)) add("header", `'${result.tcId}' is not a valid test case id.`);
      if (options.expectedTcId && result.tcId !== options.expectedTcId) add("header", `Header says ${result.tcId} but this script is for ${options.expectedTcId}.`);
      const base = path.basename(options.file);
      if (!base.startsWith(`${result.tcId}_`) || !base.endsWith(".spec.js")) add("filename", `File name must be ${result.tcId}_<slug>.spec.js (got ${base}).`);
    }
  }

  let nodes = 0, sawExpect = false, testCalls = 0;
  const importedPages = new Set<string>();
  walk(program, null, (node, parent) => {
    if (++nodes > MAX_NODES) return;
    switch (node.type) {
      case "ImportDeclaration": case "ExportAllDeclaration": case "ExportNamedDeclaration": {
        const sourceNode = node.source as Ast | null | undefined;
        if (!sourceNode) break;
        const specifier = String(sourceNode.value);
        if (specifier === ALLOWED_PACKAGE) break;
        const local = resolveLocalImport(specifier, options.file, options.automationDir);
        if (!local) { add("import", `Import '${specifier}' is not allowed. Use '${ALLOWED_PACKAGE}', '../_harness/...' or '../pages/...'.`, node); break; }
        if (!local.target.endsWith(".js")) { add("import", `Import '${specifier}' must end in .js.`, node); break; }
        if (!fs.existsSync(local.target)) { add("import", `Import '${specifier}' does not exist.`, node); break; }
        if (!local.trusted) importedPages.add(local.target);
        break;
      }
      case "ImportExpression": add("import", "Dynamic import() is not allowed.", node); break;
      case "MetaProperty": add("import", "import.meta is not allowed.", node); break;
      case "DebuggerStatement": add("debugger", "Remove debugger statements.", node); break;
      case "WithStatement": add("with", "with statements are not allowed.", node); break;
      case "Identifier": {
        if (!isReference(node, parent)) break;
        const reason = BANNED_IDENTIFIERS.get(node.name);
        if (reason) add("identifier", `'${node.name}' is not allowed: ${reason}.`, node);
        if (node.name === "expect") sawExpect = true;
        break;
      }
      case "MemberExpression": {
        const name = memberName(node);
        if (name) {
          const reason = BANNED_MEMBERS.get(name);
          if (reason) add("member", `'.${name}' is not allowed: ${reason}.`, node);
          if (name === "expect") sawExpect = true;
        } else if (node.computed && !allowedComputed(node.property)) {
          add("computed-access", "Computed property access with an expression is not allowed; use a plain property name or an index.", node);
        }
        break;
      }
      case "Property": case "PropertyDefinition": case "MethodDefinition": {
        // Object/class keys are banned too, so `{ route: ... }` cannot smuggle a name past the member check.
        const key = node.key as Ast | undefined;
        // `constructor(page) {}` in a class body is how page objects are written; only the access `x.constructor` is blocked.
        if (node.type === "MethodDefinition" && node.kind === "constructor") break;
        const keyName = !node.computed && key ? (key.type === "Identifier" ? key.name : typeof key.value === "string" ? key.value : null) : null;
        if (keyName === "__proto__" || keyName === "constructor") add("member", `'${keyName}' is not allowed as a property name.`, node);
        break;
      }
      case "ObjectPattern": {
        // Fixture destructuring such as `async ({ page, request }) => ...`.
        for (const property of (node.properties as Ast[])) {
          const name = property.type === "Property" && !property.computed ? (property.key.name ?? property.key.value) : null;
          if (name === "request") add("fixture", "The 'request' fixture is not allowed: it calls the network from Node, outside the browser allowlist.", property);
          if (name === "browser") add("fixture", "The 'browser' fixture is not allowed; use 'page' or 'context'.", property);
        }
        break;
      }
      case "Literal": case "TemplateLiteral": {
        const value = stringValue(node);
        if (value === null) break;
        for (const [pattern, label] of SECRET_LITERALS) if (pattern.test(value)) { add("secret", `String looks like ${label}. Never put credentials in scripts.`, node); break; }
        for (const url of value.match(/(?:https?:)?\/\/[^\s"'`\\)]+/gi) ?? []) {
          if (/^\/\/[^/]*$/.test(url) && !/^\/\/[\w.-]+\.[a-z]{2,}/i.test(url)) continue;
          let origin: string;
          try { origin = new URL(url.startsWith("//") ? `https:${url}` : url).origin; } catch { continue; }
          if (!options.allowedOrigins.includes(origin)) add("origin", `URL ${origin} is not in allowedOrigins. Use a path such as '/web' (baseURL is set by the harness).`, node);
        }
        break;
      }
      case "CallExpression": {
        const callee = node.callee as Ast;
        if (callee.type === "Identifier" && callee.name === "test") {
          testCalls++;
          const title = stringValue(node.arguments?.[0]);
          const id = result.tcId;
          if (title === null) add("title", "Give each test() a literal title starting with its TC id, e.g. test('[TC_001] ...').", node);
          else if (id && !title.startsWith(`[${id}]`)) add("title", `Test title must start with [${id}] (got '${title.slice(0, 40)}').`, node);
        }
        if (callee.type === "Identifier" && callee.name === "expect") sawExpect = true;
        if (callee.type === "MemberExpression") {
          const method = memberName(callee);
          if (method === "only") add("only", ".only skips every other test; remove it.", node);
          if ((method === "fill" || method === "type" || method === "pressSequentially") && callee.object) {
            const receiver = text.slice(callee.object.start, callee.object.end);
            const typed = node.arguments?.[0] as Ast | undefined;
            const literal = typed && (typed.type === "Literal" || (typed.type === "TemplateLiteral" && typed.expressions.length === 0));
            if (literal && PASSWORD_RECEIVER.test(receiver)) {
              add("password-literal", "Do not type a literal into a password field. Sign-in is handled by the saved session; for a wrong-password test use invalidPassword() from '../_harness/fixtures.js'.", node);
            }
          }
          if (method === "use" && callee.object.type === "Identifier" && callee.object.name === "test") {
            const arg = node.arguments?.[0] as Ast | undefined;
            if (!arg || arg.type !== "ObjectExpression") add("test-use", "test.use() takes an object literal.", node);
            else for (const property of arg.properties as Ast[]) {
              const name = property.type === "Property" && !property.computed ? (property.key.name ?? property.key.value) : null;
              if (!name || !ALLOWED_USE_KEYS.has(name)) add("test-use", `test.use({ ${name ?? "..."} }) is not allowed. Allowed: ${[...ALLOWED_USE_KEYS].join(", ")}.`, property);
            }
          }
        }
        break;
      }
      default: break;
    }
  });
  if (nodes > MAX_NODES) add("size", "Script is too complex to review; split it into page objects.");

  if (options.kind === "test") {
    if (testCalls === 0) add("no-test", "The file must contain at least one test('[<id>] ...', async ({ page }) => { ... }).");
    if (!sawExpect) add("no-expect", "Add at least one expect(...) assertion; a test without assertions can never fail.");
  }

  // Page objects are validated recursively; the combined list becomes part of the approval hash.
  const deps = new Set<string>();
  for (const page of importedPages) {
    if (seen.has(page)) { deps.add(posixRelative(page)); continue; }
    seen.add(page);
    let nested: ValidationResult;
    try {
      nested = validateSource(fs.readFileSync(page, "utf8"), { ...options, kind: "page", file: page, expectedTcId: undefined }, seen);
    } catch { add("import", `Cannot read imported page ${posixRelative(page)}.`); continue; }
    deps.add(posixRelative(page));
    for (const dep of nested.deps) deps.add(dep);
    for (const violation of nested.violations) violations.push({ ...violation, message: `${posixRelative(page)}: ${violation.message}` });
  }
  result.deps = [...deps].sort();
  result.ok = violations.length === 0;
  return result;
}

export function validateFile(file: string, options: Omit<ValidateOptions, "file">): ValidationResult {
  const absolute = path.resolve(file);
  return validateSource(fs.readFileSync(absolute, "utf8"), { ...options, file: absolute });
}

/** Hash of a script plus every page object it imports: the value an approval is bound to. */
export function contentHash(projectDir: string, file: string, deps: string[]): string {
  const parts = [file, ...deps].map(relative => {
    const absolute = path.resolve(projectDir, relative);
    if (!isInside(projectDir, absolute)) throw new QaError(`'${relative}' is outside the project.`, "invalid_path");
    return `${relative}\n${sha256Text(fs.readFileSync(absolute, "utf8"))}`;
  });
  return sha256Text(parts.join("\n---\n"));
}

/* -------------------------------------------------------------------------- */
/* Assertions: what a script promises to check                                 */
/* -------------------------------------------------------------------------- */

export interface Assertion {
  /** `toHaveText('Sales Order')`, `not.toBeVisible()`: the check itself, without what it is applied to. */
  signature: string;
  /** The thing being checked, such as `page.locator('#a')`. A repair may change this freely. */
  subject: string;
}

/**
 * Every expect(...).matcher(...) in a script. The signature leaves out the subject on purpose: repairing a locator
 * changes the subject but must keep the check (the matcher and the value it expects).
 */
export function extractAssertions(source: string): Assertion[] {
  const text = source.replace(/^﻿/, "");
  let program: Ast;
  try { program = parse(text, { ecmaVersion: "latest", sourceType: "module" }) as Ast; } catch { return []; }
  const squash = (node: Ast) => text.slice(node.start, node.end).split(/\s+/).filter(Boolean).join(" ");
  const found: Assertion[] = [];
  walk(program, null, (node) => {
    if (node.type !== "CallExpression" || node.callee?.type !== "MemberExpression") return;
    const matcher = memberName(node.callee);
    if (!matcher) return;
    let base: Ast = node.callee.object;
    let negated = false;
    while (base?.type === "MemberExpression") {
      if (memberName(base) === "not") negated = true;
      base = base.object;
    }
    if (base?.type !== "CallExpression") return;
    const callee = base.callee as Ast;
    const isExpect = (callee.type === "Identifier" && callee.name === "expect")
      || (callee.type === "MemberExpression" && callee.object?.type === "Identifier" && callee.object.name === "expect");
    if (!isExpect) return;
    found.push({
      signature: `${negated ? "not." : ""}${matcher}(${(node.arguments as Ast[]).map(squash).join(", ")})`,
      subject: (base.arguments as Ast[]).map(squash).join(", "),
    });
  });
  return found;
}

export interface AssertionComparison { oldCount: number; newCount: number; removed: string[]; added: string[] }

/** Which checks of the old script are gone from the new one, and which are new. Order and subjects do not matter. */
export function compareAssertions(oldSource: string, newSource: string): AssertionComparison {
  const before = extractAssertions(oldSource), after = extractAssertions(newSource);
  const pool = new Map<string, number>();
  for (const item of after) pool.set(item.signature, (pool.get(item.signature) ?? 0) + 1);
  const removed: string[] = [];
  for (const item of before) {
    const left = pool.get(item.signature) ?? 0;
    if (left > 0) pool.set(item.signature, left - 1); else removed.push(item.signature);
  }
  const added = [...pool].flatMap(([signature, count]) => Array<string>(count).fill(signature));
  return { oldCount: before.length, newCount: after.length, removed, added };
}
