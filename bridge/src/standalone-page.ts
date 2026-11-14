/**
 * The task pane for a plain browser tab (`/index.html`). Inside Excel the pane gets Office.js from Microsoft's
 * servers and works on the open workbook; in a tab there is no Excel, so the page gets a small stand-in instead.
 * Everything else is the very same page and script, served by the same bridge, with the real providers: chat,
 * settings, sign-in, history and the assistant's reasoning all work. Reading or changing cells does not:
 * `Excel.run` fails with a message the assistant can pass on.
 */

const OFFICE_SCRIPT = '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>';

const STAND_IN = `<script>
window.__EXCEL_PARTNER_STANDALONE__ = true;
window.Excel = { run: function () { return Promise.reject(new Error("This chat is open in a browser, not in Excel, so no workbook is connected. Open Excel Partner from Excel to read or change cells.")); } };
window.Office = {
  HostType: { Excel: "Excel" },
  onReady: function (callback) { setTimeout(function () { callback({ host: "Excel", platform: "Browser" }); }, 0); return Promise.resolve({ host: "Excel" }); },
  context: {
    platform: "Browser",
    document: { settings: { get: function () { return null; }, set: function () {}, saveAsync: function () {} }, addHandlerAsync: function () {}, removeHandlerAsync: function () {} },
    requirements: { isSetSupported: function () { return true; } },
    ui: {},
  },
  EventType: {},
  AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
};
</script>`;

/** `html` is `frontend/taskpane/index.html`. Its relative links keep pointing at the add-in folder through <base>. */
export function standalonePage(html: string): string {
  if (!html.includes(OFFICE_SCRIPT)) throw new Error("frontend/taskpane/index.html no longer loads office.js in the expected way");
  return html
    .replace("<head>", '<head>\n    <base href="/addin/taskpane/" />')
    .replace(OFFICE_SCRIPT, STAND_IN)
    .replace("<title>Excel Partner</title>", "<title>Excel Partner (browser)</title>");
}
