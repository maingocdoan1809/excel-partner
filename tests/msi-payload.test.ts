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

  it("asks which parts to install (add-in, gateway, or both) and shows the result of the setup", () => {
    const wxs = fs.readFileSync("installer/msi/OfficePartner.wxs", "utf8");
    expect(wxs).toContain('Id="WixUI_FeatureTree"');
    expect(wxs).toContain('<Feature Id="AddIn"');
    expect(wxs).toContain('<Feature Id="Gateway"');
    // The gateway is offered, not forced: level 2 keeps it unchecked until the person ticks it.
    expect(wxs).toMatch(/<Feature Id="Gateway"[^>]*Level="2"/);
    // The choice reaches the setup script as registry markers, so ADDLOCAL and the wizard take the same path.
    expect(wxs).toContain('Name="AddIn"');
    expect(wxs).toContain('Name="Gateway"');
    const build = fs.readFileSync("scripts/build-msi.mjs", "utf8");
    expect(build).toContain("WixToolset.UI.wixext");
    expect(build).toContain("OfficePartner.en-us.wxl");
    const bootstrap = fs.readFileSync("installer/msi/bootstrap.ps1", "utf8");
    expect(bootstrap).toContain("'status'");
    expect(bootstrap).toContain("Get-Targets");
    expect(bootstrap).not.toContain("Show-InstallStatusLater");
    expect(bootstrap.match(/ProcessStartInfo/g)).toHaveLength(1);
  });

  it("runs setup without the QuietExec pipe or an extra launcher", () => {
    const wxs = fs.readFileSync("installer/msi/OfficePartner.wxs", "utf8");
    expect(wxs).not.toContain("run-hidden.exe");
    expect(wxs).toContain("-WindowStyle Hidden -NoProfile -NonInteractive");
    const command = '&quot;[SystemFolder]WindowsPowerShell\\v1.0\\powershell.exe&quot; -WindowStyle Hidden -NoProfile -NonInteractive -ExecutionPolicy Bypass -File &quot;[INSTALLFOLDER]installer\\bootstrap.ps1&quot; -Action install -Providers &quot;[PROVIDERS]&quot;';
    for (const action of ["Install", "InstallNoSideload"]) {
      expect(wxs).toContain(`<CustomAction Id="${action}" Directory="INSTALLFOLDER" Execute="deferred" Impersonate="yes"`);
    }
    expect(wxs).toContain(`ExeCommand="${command}"`);
    expect(wxs).toContain(`ExeCommand="${command} -NoSideload"`);
    const bootstrap = fs.readFileSync("installer/msi/bootstrap.ps1", "utf8");
    expect(bootstrap).toContain("CreateNoWindow = $true");
    expect(bootstrap).toContain('<nul >> `"$logFile`" 2>&1');
  });

  it("releases files before validation even for npm installations with no installed bootstrap", () => {
    const wxs = fs.readFileSync("installer/msi/OfficePartner.wxs", "utf8");
    expect(wxs).toMatch(/<Custom Action="StopBridge" Before="InstallValidate"\s*\/>/);
    expect(wxs).toContain('<Custom Action="SetStopBridge" Before="StopBridge"');
    expect(wxs).toContain('<CustomAction Id="StopBridge" BinaryRef="$(UtilCA)" DllEntry="WixQuietExec"');
    expect(wxs).toContain('-EncodedCommand $(StopBridgeCommand)');
    const build = fs.readFileSync("scripts/build-msi.mjs", "utf8");
    expect(build).toContain('encodedCommand("stop-processes.ps1", "Stop-ProductProcesses")');
    const bootstrap = fs.readFileSync("installer/msi/bootstrap.ps1", "utf8");
    expect(bootstrap).toContain(". (Join-Path $PSScriptRoot 'stop-processes.ps1')");
    expect(bootstrap).not.toContain("function Stop-ProductProcesses");
    const payload = fs.readFileSync("scripts/build-msi-payload.mjs", "utf8");
    expect(payload).toContain('["bootstrap.ps1", "stop-processes.ps1"]');
  });

  it("purges user data after a successful full uninstall only when requested", () => {
    const wxs = fs.readFileSync("installer/msi/OfficePartner.wxs", "utf8");
    expect(wxs).toContain('<Property Id="PURGEDATA" Secure="yes" />');
    expect(wxs).toContain('-EncodedCommand $(PurgeDataCommand)');
    expect(wxs).toContain('-EncodedCommand $(PurgeGatewayCommand)');
    expect(wxs).toContain('<?define PurgeCondition = "Installed AND NOT WIX_UPGRADE_DETECTED AND NOT UPGRADINGPRODUCTCODE" ?>');
    const guard = '$(PurgeCondition) AND REMOVE=&quot;ALL&quot; AND PURGEDATA=&quot;1&quot;';
    expect(wxs).toContain(`<Custom Action="SetPurgeData" After="InstallFinalize" Condition="${guard} AND NOT (!Gateway = 3)"`);
    expect(wxs).toContain(`<Custom Action="PurgeData" After="SetPurgeData" Condition="${guard} AND NOT (!Gateway = 3)"`);
    expect(wxs).toContain(`<Custom Action="SetPurgeDataWithGateway" After="PurgeData" Condition="${guard} AND !Gateway = 3"`);
    expect(wxs).toContain(`<Custom Action="PurgeDataWithGateway" After="SetPurgeDataWithGateway" Condition="${guard} AND !Gateway = 3"`);
  });
});
