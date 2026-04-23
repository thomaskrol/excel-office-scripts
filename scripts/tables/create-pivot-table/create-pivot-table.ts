/**
 * Creates a pivot table from an existing table with specified row and value aggregations.
 * 
 * @param tableName Name of the source table for the pivot table.
 * @param location Where to place the pivot table: "New sheet" creates a new worksheet, "Existing sheet" places it below the source table unless a sheet name is specified.
 * @param rowsColumn Column name to use for pivot table rows.
 * @param valuesColumns Array of column names to aggregate in the pivot table values area.
 * @param valuesOperation Aggregation function to apply to the values columns.
 * @param sheetName The name of the sheet the pivot table should be placed on when location is Existing sheet (defaults to same sheet as table). If location is New sheet, this is the name the new sheet should have.
 * @param pivotTableName Optional name for the pivot table (auto-generates if blank or already exists)
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  location: "New sheet" | "Existing sheet" = "New sheet",
  rowsColumn: string,
  valuesColumns: Array<string>,
  valuesOperation: "Sum" | "Count" | "Average" | "Product" | "Max" | "Min" = "Sum",
  columnsColumn?: string,
  sheetName?: string,
  pivotTableName?: string
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table '${tableName}' not found.`);
  }

  if (table.getRowCount() === 0) {
    throw new Error(`Table '${tableName}' has no data.`);
  }

  // make sure specified columns exist
  const tableCols = table.getColumns().map((col) => col.getName());
  if (!tableCols.includes(rowsColumn)) {
    throw new Error(`There is no column '${rowsColumn}' in table '${tableName}'.`);
  }
  else {
    valuesColumns.forEach((colName) => {
      if (!tableCols.includes(colName)) {
        throw new Error(`There is no column '${colName}' in table '${tableName}'.`);
      }
    });
  }

  // validate operation
  const operation = ExcelScript.AggregationFunction[valuesOperation.toLowerCase() as keyof typeof ExcelScript.AggregationFunction];
  if (!operation) {
    throw new Error(`Invalid operation: ${valuesOperation}`);
  }

  // get range of where to add pivot table
  let locationRange: ExcelScript.Range;
  if (location === "New sheet") {
    locationRange = workbook.addWorksheet(sheetName).getRange("A1");
  } else {
    let locationReference: ExcelScript.Worksheet;
    if (sheetName) {
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        throw new Error(`There is no worksheet "${sheetName}" in the Excel file.`);
      }

      locationReference = sheet;
    } else {
      locationReference = table.getWorksheet();
    }
    const lastUsedRow = locationReference.getUsedRange().getLastRow();
    // 2 for offset + 1 for 0-based index
    console.log(`Destination: '${locationReference.getName()}'!A${lastUsedRow.getRowIndex() + 3}`);
    locationRange = lastUsedRow.getCell(0, 0).getOffsetRange(2, 0);
  }

  const usedSheetName = locationRange.getWorksheet().getName();

  // get next available pivot table name if provided one is taken or blank
  const existingPivotTables = workbook.getPivotTables().map((pvtTbl) => pvtTbl.getName());
  if (!pivotTableName || existingPivotTables.includes(pivotTableName)) {
    const defaultName = "PivotTable";
    let i = 1;
    const maxAttempts = 100;
    while (existingPivotTables.includes(defaultName + i) && i < maxAttempts) {
      i++;
    }

    if (i >= maxAttempts) {
      throw new Error(`Unable to generate unique pivot table name after ${i} attempts`);
    }

    pivotTableName = defaultName + i;
  }

  const pivotTable = workbook.addPivotTable(pivotTableName, table, locationRange);

  // add rows field to pivot table
  pivotTable.addRowHierarchy(pivotTable.getHierarchy(rowsColumn));

  // add columns field to pivot table if it exists
  if (columnsColumn) {
    pivotTable.addColumnHierarchy(pivotTable.getHierarchy(columnsColumn));
  }

  // add values fields to pivot table
  valuesColumns.forEach((colName) => {
    const valuesField = pivotTable.addDataHierarchy(pivotTable.getHierarchy(colName));
    valuesField.setSummarizeBy(operation);
  });

  return {
    "message": "Successfully created a pivot table.",
    "createdPivotTableName": pivotTableName,
    "usedSheetName": usedSheetName
  }
}