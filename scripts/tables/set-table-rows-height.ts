/**
 * Sets the row height for all rows in a table.
 * 
 * @param tableName Name of the table to modify
 * @param rowHeight Height in points to apply to all table rows (defaults to 14.25)
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  rowHeight?: number
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table '${tableName}' not found.`);
  }

  // use default row height unless provided
  table.getRange().getFormat().setRowHeight(rowHeight || 14.25);
}