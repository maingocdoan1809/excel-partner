import type { HostId } from "../../../protocol/src/index.js";

/** Largest file a backup copy is made of. */
export const MAX_BACKUP_BYTES = 100 * 1024 * 1024;

const APP_NAMES: Record<HostId, string> = { excel: "Excel", word: "Word", powerpoint: "PowerPoint" };

/** The open file as zip bytes (xlsx, docx or pptx), read through Office in 4 MB slices. */
export function readOpenFile(host: HostId): Promise<Uint8Array> {
  const app = APP_NAMES[host];
  return new Promise((resolve, reject) => {
    try {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (opened) => {
        if (opened.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(opened.error?.message ?? `${app} would not hand over the file`));
          return;
        }
        const file = opened.value;
        if (file.size > MAX_BACKUP_BYTES) {
          file.closeAsync();
          reject(new Error(`the file is ${Math.round(file.size / 1024 / 1024)} MB, over the ${MAX_BACKUP_BYTES / 1024 / 1024} MB backup limit`));
          return;
        }
        const slices: Uint8Array[] = new Array(file.sliceCount);
        let remaining = file.sliceCount;
        let done = false;
        const finish = (error?: Error) => {
          if (done) return;
          done = true;
          file.closeAsync();
          if (error) return reject(error);
          const bytes = new Uint8Array(slices.reduce((sum, slice) => sum + slice.length, 0));
          let at = 0;
          for (const slice of slices) {
            bytes.set(slice, at);
            at += slice.length;
          }
          resolve(bytes);
        };
        if (remaining === 0) return finish(new Error("the file is empty"));
        for (let index = 0; index < file.sliceCount; index++) {
          file.getSliceAsync(index, (slice) => {
            if (slice.status !== Office.AsyncResultStatus.Succeeded) return finish(new Error(slice.error?.message ?? "could not read part of the file"));
            slices[index] = Uint8Array.from(slice.value.data as ArrayLike<number>);
            if (--remaining === 0) finish();
          });
        }
      });
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
}
