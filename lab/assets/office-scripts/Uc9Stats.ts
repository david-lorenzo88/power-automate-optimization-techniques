// UC9 checksum: "<rows>|<amount total>" over the `Orders` table (same call for SLOW and FAST).
// Blank rows (a reset table keeps one) are ignored.
function main(workbook: ExcelScript.Workbook): string {
  const table = workbook.getTable("Orders");
  let n = 0;
  let sum = 0;
  if (table.getRowCount() > 0) {
    for (const row of table.getRangeBetweenHeaderAndTotal().getValues()) {
      if (String(row[0]) === "") continue;
      n++;
      sum += Number(row[2]) || 0;
    }
  }
  return `${n}|${sum.toFixed(2)}`;
}
