/** Extensions Excel adds to a workbook's name; an unsaved workbook ("Book2") has none. */
const WORKBOOK_EXTENSION = /\.(?:xlsx|xlsm|xlsb|xls|xltx|xltm|csv)$/i;

/** The name as people say it: without the file extension. */
export function workbookStem(name: string): string {
  return name.trim().replace(WORKBOOK_EXTENSION, "");
}

/** True when two names mean the same workbook: "Book2" and "book2.xlsx" do, "Book2" and "Book20" do not. */
export function sameWorkbookName(a: string, b: string): boolean {
  const left = workbookStem(a).toLowerCase();
  return left !== "" && left === workbookStem(b).toLowerCase();
}
