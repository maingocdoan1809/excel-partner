/// <reference lib="dom" />

let remembered = "";

/** The open file's url when Office already reports it, else the last one found by `currentDocumentUrl`; "" for a file that is not saved. */
export function knownDocumentUrl(): string {
  try {
    return Office.context.document.url || remembered;
  } catch {
    return remembered;
  }
}

/** The open file's url; asks Office for the file properties when the url is not at hand yet. "" for a file that is not saved. */
export async function currentDocumentUrl(): Promise<string> {
  const direct = knownDocumentUrl();
  if (direct) return direct;
  return new Promise((resolve) => {
    const timer = window.setTimeout(() => resolve(""), 2500);
    try {
      Office.context.document.getFilePropertiesAsync((result) => {
        window.clearTimeout(timer);
        remembered = result.status === Office.AsyncResultStatus.Succeeded ? result.value.url ?? "" : "";
        resolve(remembered);
      });
    } catch { window.clearTimeout(timer); resolve(""); }
  });
}
