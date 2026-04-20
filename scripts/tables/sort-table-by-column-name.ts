/**
 * Sorts a table in ascending order by the specified column.
 *
 * @param tableName Name of the table to sort.
 * @param columnName Name of the column to sort by.
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  columnName: string
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table '${tableName}' not found.`);
  }

  // Get the column index based on the column name
  const columnIndex = table.getHeaderRowRange().getValues()[0].indexOf(columnName);

  if (columnIndex === -1) {
    throw new Error(`Column '${columnName}' not found.`);
  }

  table.getSort().apply([{ key: columnIndex, ascending: true }]);
}