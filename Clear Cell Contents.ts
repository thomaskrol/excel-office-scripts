function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  columnName: string,
  rowIndex: number
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table "${tableName}" not found.`);
  }

  // Get all headers and find the column index
  const headers = table.getHeaderRowRange().getValues()[0];
  const columnIndex = headers.indexOf(columnName);
  if (columnIndex === -1) {
    throw new Error(`Column "${columnName}" not found.`);
  }

  // Get the row (excluding the header row);
  const rowRange = table.getRangeBetweenHeaderAndTotal();
  const targetCell = rowRange.getCell(rowIndex, columnIndex);
  targetCell.clear(ExcelScript.ClearApplyTo.contents);
}