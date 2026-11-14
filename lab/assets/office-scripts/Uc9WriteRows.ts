// UC9 FAST: append all rows to the `Orders` table with ONE setValues() call.
// rowsJson: JSON array of [OrderId, Region, Amount] arrays. Returns the table's data row count.
function main(workbook: ExcelScript.Workbook, rowsJson: string): number {
  const rows: (string | number)[][] = JSON.parse(rowsJson);
  const table = workbook.getTable("Orders");
  if (rows.length === 0) return table.getRowCount();
  const sheet = table.getWorksheet();
  const header = table.getHeaderRowRange();
  const top = header.getRowIndex();
  const left = header.getColumnIndex();
  // A freshly reset table keeps one blank row: overwrite it instead of appending after it.
  let existing = table.getRowCount();
  if (existing === 1 && String(table.getRangeBetweenHeaderAndTotal().getValues()[0][0]) === "") existing = 0;
  sheet.getRangeByIndexes(top + 1 + existing, left, rows.length, 3).setValues(rows);
  table.resize(sheet.getRangeByIndexes(top, left, 1 + existing + rows.length, 3));
  return existing + rows.length;
}
