/**
 * Auto-fits the column widths for all columns in a table to best fit their contents.
 *
 * @param tableName Name of the table to auto-fit columns for.
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table "${tableName}" not found.`);
  }

  table.getRange().getFormat().autofitColumns();
}