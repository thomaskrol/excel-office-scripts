/**
 * Converts values in specified table columns to Proper Case if they are entirely uppercase or entirely lowercase.
 * Handles accented characters common in European names.
 *
 * @param tableName Name of the table to process.
 * @param columnsToFix Array of column names whose values should be converted to Proper Case.
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  columnsToFix: string[];
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table "${tableName}" not found.`);
  }

  // Process each column
  columnsToFix.forEach(columnName => {
    let column = table.getColumnByName(columnName);
    let values = column.getRange().getValues();

    // Loop through each row in the column
    for (let i = 0; i < values.length; i++) {
      let cellValue = values[i][0];
      if (typeof cellValue !== "string" || cellValue.trim() === "") {
        continue;
      }

      // Check if the value is all uppercase or all lowercase
      if (cellValue === cellValue.toUpperCase() || cellValue === cellValue.toLowerCase()) {
        // Convert to Proper Case while handling accents
        let properCaseValue = cellValue.toLowerCase().replace(
          /(^|\\s)([a-záéíóúüñâàäêëîïôöûüç])/g,
          (_, boundary, letter) => boundary + letter.toUpperCase();
        );

        // Update the value in the array
        values[i][0] = properCaseValue
      }
    }

    // Write the updated values back to the column
    column.getRange().setValues(values);
  });
}