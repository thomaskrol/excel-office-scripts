function main(workbook: ExcelScript.Workbook) {
  // Set workbook calculation mode
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);
}