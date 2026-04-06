function main(
  workbook: ExcelScript.Workbook,
  sheetName: string
) {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`Worksheet ${sheetName} not found.`);
  }

  sheet.delete();
}