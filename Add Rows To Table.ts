function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  inputJson: string
) {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table "${tableName}" not found.`);
  }

  const columnNames = table.getColumns().map(col => col.getName().trim());
  let inputData: Record<string, string>[];
  try {
    inputData = JSON.parse(inputJson);
  } catch (e) {
    throw new Error(`Failed to parse input JSON: ${e.message}`);
  }

  if (!Array.isArray(inputData) || inputData.length === 0) {
    return "Input JSON contains no rows to add.";
  }

  // validate input keys
  const inputKeys = Object.keys(inputData[0]).map(k => k.trim());
  for (const key of inputKeys) {
    if (!columnNames.includes(key)) {
      throw new Error(`Input key "${key}" does not match any column in the table "${tableName}".`);
    }
  }

  // build 2D array matching table column order
  const rows = inputData.map(obj => columnNames.map(colName => obj[colName] ?? undefined));
  console.log(rows);
  table.addRows(-1, rows);
  return "Rows added successfully.";
}