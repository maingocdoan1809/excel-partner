import { afterEach, describe, expect, it } from "vitest";
import http from "node:http";
import { Bridge } from "../bridge/src/server.js";
import { loadConfig } from "../bridge/src/config.js";
import { standalonePage } from "../bridge/src/standalone-page.js";
import { MockAdapter } from "../providers/src/mock.js";
import { ProviderRegistry } from "../providers/src/registry.js";
import type { ConnectedFile } from "../protocol/src/index.js";
import { PaneClient, getFreePort } from "./helpers.js";

let bridge: Bridge | null = null;
afterEach(async () => { await bridge?.close(); bridge = null; });

async function start() {
  const port = await getFreePort();
  const config = loadConfig({ host: "127.0.0.1", port, tls: null, publicBaseUrl: `http://127.0.0.1:${port}`, logLevel: "error" });
  bridge = new Bridge(config, { usageFile: null, registry: new ProviderRegistry().register(new MockAdapter({ id: "claude", reply: "ok" })) });
  return (await bridge.listen()).url;
}

/** `fetch` forbids overriding Host, so use http.request to exercise rebinding-style requests. */
function request(port: number, pathname: string, headers: Record<string, string>) {
  return new Promise<{ status: number }>((resolve, reject) => {
    const req = http.request({ host: "127.0.0.1", port, path: pathname, headers }, (res) => {
      res.resume();
      res.on("end", () => resolve({ status: res.statusCode ?? 0 }));
    });
    req.on("error", reject);
    req.end();
  });
}

const OFFICE_CDN = "appsforoffice.microsoft.com";

describe("the task pane in a browser tab", () => {
  it("is served at / and /index.html: the add-in's page, with a stand-in for Office.js and its own links intact", async () => {
    const url = await start();
    const real = await (await fetch(`${url}/addin/taskpane/index.html`)).text();
    expect(real).toContain(OFFICE_CDN);

    for (const route of ["/", "/index.html"]) {
      const response = await fetch(url + route);
      expect(response.status, route).toBe(200);
      expect(response.headers.get("content-type")).toContain("text/html");
      expect(response.headers.get("cache-control")).toBe("no-store");
      const page = await response.text();
      expect(page).not.toContain(OFFICE_CDN);
      expect(page).toContain("__EXCEL_PARTNER_STANDALONE__");
      expect(page).toContain('<base href="/addin/taskpane/" />');
      expect(page).toContain("../dist/taskpane.js");
      expect(page).not.toContain("__ADDIN_BUILD_ID__");
      expect(page).toContain("<title>Excel Partner (browser)</title>");
    }
    // The links the page relies on resolve once <base> points at the add-in folder.
    for (const asset of ["/addin/taskpane/styles.css", "/addin/dist/taskpane.js"]) {
      const response = await fetch(url + asset);
      await response.arrayBuffer(); // read it all: an open response would keep the bridge from closing
      expect(response.status, asset).toBe(200);
    }
  });

  it("leaves the page Excel loads untouched", async () => {
    const url = await start();
    const page = await (await fetch(`${url}/addin/taskpane/index.html`)).text();
    expect(page).toContain(OFFICE_CDN);
    expect(page).not.toContain("__EXCEL_PARTNER_STANDALONE__");
    expect(page).not.toContain("<base");
  });

  it("is refused for a page from another origin or a rebound host, like the rest of the bridge", async () => {
    const url = await start();
    const port = Number(new URL(url).port);
    for (const route of ["/", "/index.html"]) {
      expect((await request(port, route, {})).status, route).toBe(200);
      expect((await request(port, route, { origin: "https://evil.example" })).status, route).toBe(403);
      expect((await request(port, route, { host: `evil.example:${port}` })).status, route).toBe(403);
    }
  });

  it("refuses to build the page when the add-in page no longer loads office.js as expected", () => {
    expect(() => standalonePage("<html><head></head></html>")).toThrow("office.js");
  });

  it("is not a file another workbook can see or send a request to, but sees the real ones itself", async () => {
    const url = await start();
    const files = (pane: PaneClient): ConnectedFile[] => {
      const last = pane.received.filter((m) => m.type === "connected_files").at(-1);
      return last && last.type === "connected_files" ? last.files : [];
    };
    const excel = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    excel.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "excel-connection-1" });
    await excel.waitFor((m) => m.type === "hello_ok");
    excel.send({ type: "workbook_info", name: "Budget.xlsx" });
    await excel.waitFor((m) => m.type === "connected_files" && m.files[0]?.name === "Budget.xlsx");

    const tab = await PaneClient.connect(`${url.replace("http", "ws")}/ws`);
    tab.send({ type: "hello", protocolVersion: 1, workbookConnectionId: "browser-tab-0001", clientInfo: { platform: "Browser", standalone: true } });
    await tab.waitFor((m) => m.type === "hello_ok");
    await tab.waitFor((m) => m.type === "connected_files" && m.files.some((f) => f.name === "Budget.xlsx"));

    expect(files(tab).map((f) => [f.name, f.current])).toEqual([["", true], ["Budget.xlsx", false]]);
    // Give the bridge time to announce the new connection to everyone; Excel's list must not have changed.
    await new Promise((resolve) => setTimeout(resolve, 150));
    expect(files(excel)).toEqual([{ id: "f1", name: "Budget.xlsx", host: "excel", current: true }]);

    await tab.close();
    await excel.close();
  });
});
