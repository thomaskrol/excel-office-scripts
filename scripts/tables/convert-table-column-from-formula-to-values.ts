/**
 * Converts a table column's formulas to static values by copying the range as values-only.
 * Clears any active filters before converting to ensure all rows are affected.
 *
 * @param tableName Name of the table containing the column.
 * @param columnName Name of the column to convert.
 * @param fullColumn If true, converts the entire column. If false, only the last n rows are converted.
 * @param numberOfRowsFromEnd Number of rows from the end of the column to convert (required when fullColumn is false).
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  columnName: string,
  fullColumn: boolean = false,
  numberOfRowsFromEnd?: number
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table '${tableName}' not found.`);
  }

  const column = table.getColumnByName(columnName);
  if (!column) {
    throw new Error(`Column '${columnName}' not found.`);
  }

  // clear filters so following actions are not impacted
  table.getAutoFilter().clearCriteria();

  const totalColRange = column.getRangeBetweenHeaderAndTotal();

  let range: ExcelScript.Range;
  if (fullColumn) {
    range = totalColRange;

  } else if (numberOfRowsFromEnd) {
    range = totalColRange.getLastCell().getOffsetRange(1 - numberOfRowsFromEnd, 0).getAbsoluteResizedRange(numberOfRowsFromEnd, 1);

  } else {
    throw new Error(`Parameter 'numberOfRowsFromEnd' is required when fullColumn is set to false.`);
  }

  range.copyFrom(
    range,
    ExcelScript.RangeCopyType.values,
    false,
    false
  );
}