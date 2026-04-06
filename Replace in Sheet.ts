function main(
  workbook: ExcelScript.Workbook,
  sheetName: string,
  oldValue: string,
  newValue: string,
  matchCase: boolean = false,
  matchEntireCellContents: boolean = false
) {
  const sheet: ExcelScript.Worksheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`Worksheet "${sheetName}" not found.`);
  }

  sheet.replaceAll(oldValue, newValue, { matchCase: matchCase, completeMatch: matchEntireCellContents });
}