/** The open workbook as an .xlsx file (zip bytes), read through Office. */
export function readWorkbookFile(): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (opened) => {
      if (opened.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(opened.error?.message ?? "Excel would not hand over the workbook file"));
        return;
      }
      const file = opened.value;
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
      if (remaining === 0) return finish(new Error("the workbook file is empty"));
      for (let index = 0; index < file.sliceCount; index++) {
        file.getSliceAsync(index, (slice) => {
          if (slice.status !== Office.AsyncResultStatus.Succeeded) return finish(new Error(slice.error?.message ?? "could not read part of the workbook file"));
          slices[index] = Uint8Array.from(slice.value.data as ArrayLike<number>);
          if (--remaining === 0) finish();
        });
      }
    });
  });
}
