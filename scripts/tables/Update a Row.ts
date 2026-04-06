/**
 * Updates one or more column values in a table row identified by a key column value.
 * Returns the Excel row number of the updated row on success.
 *
 * @param tableName Name of the table containing the row to update.
 * @param keyColumnName Name of the column used to identify the target row.
 * @param keyValue The value in the key column that identifies the row to update.
 * @param updatesJson JSON string of column-name-to-value mappings to apply (e.g. {"Status": "Done", "Count": 5}).
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  keyColumnName: string,
  keyValue: string,
  updatesJson: string
) {
  // Parse updates from JSON string
  let updates: { [column: string]: string | number | boolean };
  try {
    updates = JSON.parse(updatesJson);
  } catch (e) {
    return {
      success: false,
      message: "Invalid JSON in updatesJson parameter"
    };
  }

  const table = workbook.getTable(tableName);
  if (!table) {
    return {
      success: false,
      message: `Table "${tableName}" not found`
    };
  }

  table.clearFilters();
  const keyCol = table.getColumnByName(keyColumnName);
  if (!keyCol) {
    return {
      success: false,
      message: `Column "${keyColumnName}" not found`
    };
  }

  // Get column range
  const colRange = keyCol.getRangeBetweenHeaderAndTotal();
  const data = colRange.getValues();

  // Get all data and find matching row
  let targetRowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i]) === String(keyValue)) {
      targetRowIndex = i
      break;
    }
  }
  console.log(`targetRowIndex: ${targetRowIndex}`);
  if (targetRowIndex === -1) {
    return {
      success: false,
      message: `No row found with ${keyColumnName} = "${keyValue}"`
    }
  }

  // Update each specified column
  let rowCell: ExcelScript.Range
  for (const [column, value] of Object.entries(updates)) {
    const colIndex = table.getColumnByName(column).getIndex();

    if (colIndex === -1) {
      return {
        success: false,
        message: `Column "${keyColumnName}" not found`
      }
    }

    rowCell = table.getColumnByName(column).getRangeBetweenHeaderAndTotal().getCell(targetRowIndex, 0);
    rowCell.setValue(value);
  }

  return {
    success: true,
    message: "Row updated successfully",
    // +1 for header, +1 for Excel 1-indexing
    row: targetRowIndex + table.getHeaderRowRange().getRowIndex() + 2
  };
}