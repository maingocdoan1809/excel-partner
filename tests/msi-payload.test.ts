import fs from "node:fs";
import { describe, expect, it } from "vitest";
// @ts-expect-error Build tooling ships as plain JavaScript, like the bin tools.
import { commandShim, msiVersion, pruneReason, readPins } from "../scripts/msi/payload.mjs";

const manifest = JSON.parse(fs.readFileSync("bin/runtime-manifest.json", "utf8"));
const pkg = JSON.parse(fs.readFileSync("package.json", "utf8")) as { version: string; dependencies: Record<string, string>; optionalDependencies: Record<string, string>; engines: { node: string } };
const clone = () => JSON.parse(JSON.stringify(manifest));

describe("MSI pins", () => {
  it("pin the Node.js, SQLite binary and Copilot CLI of every architecture to exact versions and hashes", () => {
    for (const arch of ["x64", "arm64"] as const) {
      const pins = readPins(manifest, arch);
      expect(pins.node.url).toBe(`https://nodejs.org/dist/v${pins.node.version}/node-v${pins.node.version}-win-${arch}.zip`);
      expect(pins.sqlite.url).toContain(`/v${pins.sqlite.version}/better-sqlite3-v${pins.sqlite.version}-node-v${pins.node.abi}-win32-${arch}.tar.gz`);
      expect(pins.copilot).toEqual({ package: "@github/copilot", version: manifest.providers.copilot.version, sdkVersion: manifest.providers.copilot.sdkVersion });
    }
  });

  it("keep the SQLite binary on the same version as the driver the package depends on", () => {
    expect(readPins(manifest, "x64").sqlite.version).toBe(pkg.optionalDependencies["better-sqlite3"]);
    expect(manifest.providers.copilot.sdkVersion).toBe(pkg.dependencies["@github/copilot-sdk"]);
  });

  it("name a Node.js release the package supports, whose ABI matches the SQLite binary", () => {
    const [major, minor] = manifest.node.version.split(".").map(Number) as [number, number];
    expect(pkg.engines.node).toContain(">=22.12.0 <26.0.0");
    expect(major === 22 && minor >= 12).toBe(true);
    expect(manifest.node.abi).toBe(127);
  });

  it("are the same Node.js release the stand-alone installer script downloads", () => {
    const script = fs.readFileSync("installer/install-office-partner.cmd", "utf8");
    expect(script).toContain(`{ '${manifest.node.version}' }`);
  });

  it("refuse a loose or missing pin", () => {
    const loose = clone(); loose.node.version = "22";
    expect(() => readPins(loose, "x64")).toThrow(/exact/);
    const wrongAbi = clone(); wrongAbi.node.abi = 137;
    expect(() => readPins(wrongAbi, "x64")).toThrow(/ABI/);
    const noHash = clone(); delete noHash.node.win32.arm64.sha256;
    expect(() => readPins(noHash, "arm64")).toThrow(/sha256/);
    const noSqlite = clone(); noSqlite.native["better-sqlite3"].prebuilt["win32-x64"].sha256 = "abc";
    expect(() => readPins(noSqlite, "x64")).toThrow(/SHA-256/);
    const noSdk = clone(); delete noSdk.providers.copilot.sdkVersion;
    expect(() => readPins(noSdk, "x64")).toThrow(/sdkVersion/);
    expect(() => readPins(manifest, "ia32" as never)).toThrow(/architecture/);
  });
});

describe("MSI payload pruning", () => {
  it("drops documentation, typings, source maps and the SDK's in-process runtime", () => {
    expect(pruneReason("zod/lib/index.d.ts", false)).toBeTruthy();
    expect(pruneReason("marked/lib/marked.esm.js.map", false)).toBeTruthy();
    expect(pruneReason("ws/README.md", false)).toBeTruthy();
    expect(pruneReason("ws/CHANGELOG", false)).toBeTruthy();
    expect(pruneReason("@github/copilot-sdk/docs", true)).toBeTruthy();
    expect(pruneReason("@github/copilot-sdk-win32-x64", true)).toBeTruthy();
    expect(pruneReason("better-sqlite3/deps", true)).toBeTruthy();
    expect(pruneReason("better-sqlite3/src", true)).toBeTruthy();
  });

  it("keeps everything that runs, and every licence text", () => {
    expect(pruneReason("@github/copilot-sdk", true)).toBeNull();
    expect(pruneReason("@github/copilot-sdk/dist/client.js", false)).toBeNull();
    expect(pruneReason("better-sqlite3/lib", true)).toBeNull();
    expect(pruneReason("better-sqlite3/build/Release/better_sqlite3.node", false)).toBeNull();
    expect(pruneReason("@koromix/koffi-win32-x64", true)).toBeNull();
    expect(pruneReason("ws/LICENSE", false)).toBeNull();
    expect(pruneReason("zod/LICENSE.md", false)).toBeNull();
    expect(pruneReason("some/NOTICE", false)).toBeNull();
    // Looks like documentation, is code: yaml keeps its document model in dist/doc, and a package may have a file called history.js.
    expect(pruneReason("yaml/dist/doc/directives.js", false)).toBeNull();
    expect(pruneReason("yaml/dist/doc", true)).toBeNull();
    expect(pruneReason("tool/lib/history.js", false)).toBeNull();
    // A package that happens to be called `test` or `docs` is not a folder of tests.
    expect(pruneReason("office-partner/node_modules/test", true)).toBeNull();
    expect(pruneReason("docs", true)).toBeNull();
  });
});

describe("MSI files", () => {
  it("writes a CRLF command that prefers the pinned Node.js this installer put in place", () => {
    const shim = commandShim();
    expect(shim).toContain("\r\n");
    expect(shim).not.toMatch(/[^\r]\n/);
    expect(shim).toContain("%LOCALAPPDATA%\\Programs\\nodejs\\node.exe");
    expect(shim).toContain("%dp0%\\node_modules\\office-partner\\bin\\office-partner.mjs");
  });

  it("turns a package version into an MSI version", () => {
    expect(msiVersion("0.1.26")).toBe("0.1.26");
    expect(msiVersion("1.2.3-beta.4")).toBe("1.2.3");
    expect(() => msiVersion("300.0.0")).toThrow();
  });

  it("installs where `npm install -g` would, and runs the bootstrap script the build copies into the package", () => {
    const wxs = fs.readFileSync("installer/msi/OfficePartner.wxs", "utf8");
    expect(wxs).toContain('Scope="perUser"');
    expect(wxs).toContain('Id="AppDataFolder"');
    expect(wxs).toContain('Name="npm"');
    expect(wxs).toContain("installer\\bootstrap.ps1");
    expect(fs.existsSync("installer/msi/bootstrap.ps1")).toBe(true);
    const bootstrap = fs.readFileSync("installer/msi/bootstrap.ps1", "utf8");
    expect(bootstrap).toContain("bin\\runtime-manifest.json");
    expect(bootstrap).not.toMatch(/SHASUMS/); // the hash comes from the manifest, never from the download site
  });

  it("is built with the WiX release the definition was written for", () => {
    expect(fs.readFileSync("scripts/build-msi.mjs", "utf8")).toContain('WIX_VERSION = "5.0.2"');
    expect(fs.readFileSync(".github/workflows/msi.yml", "utf8")).toContain("--version 5.0.2");
  });
});
