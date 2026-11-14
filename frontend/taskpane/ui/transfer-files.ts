const IMAGE_EXTENSIONS: Record<string, string> = {
  "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif", "image/webp": "webp", "image/bmp": "bmp",
};

type ClipboardLike = Pick<DataTransfer, "files" | "getData"> & { items?: DataTransfer["items"] };

/** Screenshots arrive as "image.png" every time; give each one its own name so the @mentions stay distinct. */
function pastedName(file: File, index: number, now: number): string {
  if (file.name && !/^image\.\w+$/i.test(file.name)) return file.name;
  const extension = IMAGE_EXTENSIONS[file.type] ?? file.name.split(".").pop() ?? "bin";
  return `pasted-image-${now}${index ? `-${index + 1}` : ""}.${extension}`;
}

type DropLike = Pick<DataTransfer, "files"> & { items?: DataTransfer["items"]; types?: DataTransfer["types"] };

/** True while the user drags files (not text) over the window. */
export function isFileDrag(data: Pick<DropLike, "types"> | null | undefined): boolean {
  return Array.from(data?.types ?? []).includes("Files");
}

/** Files dropped from Explorer, with the number of folders skipped (a folder cannot be read as a file). */
export function filesFromDrop(data: DropLike | null | undefined): { files: File[]; folders: number } {
  const items = Array.from(data?.items ?? []).filter((item) => item.kind === "file");
  if (!items.length) return { files: Array.from(data?.files ?? []), folders: 0 };
  const files: File[] = [];
  let folders = 0;
  for (const item of items) {
    if (item.webkitGetAsEntry?.()?.isDirectory) folders++;
    else { const file = item.getAsFile(); if (file) files.push(file); }
  }
  return { files, folders };
}

/**
 * Files or images that Ctrl+V should attach, or [] when the paste should stay a normal text paste.
 * Cells copied from Excel carry a picture next to their text, so any text on the clipboard wins.
 */
export function filesFromClipboard(data: ClipboardLike | null | undefined, now = Date.now()): File[] {
  if (!data || data.getData("text/plain")) return [];
  const fromItems = Array.from(data.items ?? [])
    .filter((item) => item.kind === "file")
    .flatMap((item) => item.getAsFile() ?? []);
  const files = fromItems.length ? fromItems : Array.from(data.files ?? []);
  return files.map((file, index) => {
    const name = pastedName(file, index, now);
    return name === file.name ? file : new File([file], name, { type: file.type });
  });
}
