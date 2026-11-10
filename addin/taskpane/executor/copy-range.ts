/** Resolve the whole area Office.js will write, including automatic expansion of a smaller destination. */
export async function copyDestination(context: Excel.RequestContext, source: Excel.Range, destination: Excel.Range): Promise<Excel.Range> {
  source.load("rowCount,columnCount");
  destination.load("rowCount,columnCount");
  await context.sync();
  return destination.getCell(0, 0).getResizedRange(
    Math.max(source.rowCount, destination.rowCount) - 1,
    Math.max(source.columnCount, destination.columnCount) - 1,
  );
}
