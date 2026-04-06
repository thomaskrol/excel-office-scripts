/**
 * Sets the workbook calculation mode to automatic, ensuring all formulas are recalculated.
 * Useful when running scripts after a workbook has had calculation mode set to manual.
 */
function main(workbook: ExcelScript.Workbook) {
  // Set workbook calculation mode
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);
}