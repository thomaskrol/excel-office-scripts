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