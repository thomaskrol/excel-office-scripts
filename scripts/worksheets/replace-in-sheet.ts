/**
 * Replaces all occurrences of a value in a worksheet with a new value.
 *
 * @param sheetName Name of the worksheet to perform replacements in.
 * @param oldValue The value to search for.
 * @param newValue The value to replace matches with.
 * @param matchCase Whether the search should be case-sensitive (defaults to false).
 * @param matchEntireCellContents Whether to match only cells whose entire contents equal oldValue (defaults to false).
 */
function main(
  workbook: ExcelScript.Workbook,
  sheetName: string,
  oldValue: string,
  newValue: string,
  matchCase: boolean = false,
  matchEntireCellContents: boolean = false
) {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`Worksheet "${sheetName}" not found.`);
  }

  sheet.replaceAll(oldValue, newValue, { matchCase: matchCase, completeMatch: matchEntireCellContents });
}