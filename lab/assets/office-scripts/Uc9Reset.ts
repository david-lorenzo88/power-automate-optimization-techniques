// UC9 reset: delete every data row of the `Orders` table.
function main(workbook: ExcelScript.Workbook): number {
  const table = workbook.getTable("Orders");
  const n = table.getRowCount();
  if (n > 0) table.deleteRowsAt(0, n);
  return table.getRowCount();
}
