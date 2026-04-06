/**
 * Hides a worksheet in the workbook.
 *
 * @param sheetName Name of the worksheet to hide.
 */
function main(
  workbook: ExcelScript.Workbook,
  sheetName: string
) {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`Worksheet "${sheetName}" not found.`);
  }

  // Set sheet visibility to hidden
  sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
}