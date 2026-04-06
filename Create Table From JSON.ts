/**
  * Creates a table from JSON data on a specified worksheet.
  * 
  * @param sheetName Name of the worksheet where the table will be created
  * @param startCell Top-left cell address for the table (e.g., "A1")
  * @param inputJson JSON string containing array of objects with consistent keys
  * @param tableName Optional name for the created table
  */
function main(
  workbook: ExcelScript.Workbook,
  sheetName: string,
  startCell: string,
  inputJson: string,
  tableName?: string
) {
  let returnMsg: string

  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) throw new Error(`Worksheet "${sheetName}" not found.`);

  let inputData: Record<string, string | number | boolean>[];
  try {
    inputData = JSON.parse(inputJson);
  } catch (e) {
    throw new Error(`Failed to parse input JSON: ${e.message}`);
  }

  // nothing to write
  if (!Array.isArray(inputData) || inputData.length === 0) return {
    "message": "Input JSON has no data."
  }

  // calculate range based on data dimensions
  const numCols = Object.keys(inputData[0]).length;

  // convert startCell to a full range
  const startRange = sheet.getRange(startCell);
  const endRange = startRange.getOffsetRange(0, numCols - 1);
  const tableRange = `${startCell}:${endRange.getAddress().split('!')[1]}`

  // create table with headers only
  const table = sheet.addTable(tableRange, true);
  if (tableName) {
    try {
      table.setName(tableName);
    } catch (e) {
      const allTables = workbook.getTables().map(tbl => tbl.getName());
      if (allTables.includes(tableName)) {
        returnMsg = `A table already exists with the name "${tableName}".`
      } else {
        throw new Error(e);
      }
    }
  }

  // set column headers
  const inputKeys = Object.keys(inputData[0]).map(k => k.trim());
  const columns = table.getColumns();
  inputKeys.forEach((key, index) => {
    columns[index].setName(key);
  });

  // add data rows
  const rowsData = inputData.map(obj => inputKeys.map(key => obj[key]));
  table.addRows(-1, rowsData);
  return {
    "message": "Successfully created table." + ((returnMsg) ? "\ " + returnMsg : ""), "createdTableName": table.getName()
  };
}