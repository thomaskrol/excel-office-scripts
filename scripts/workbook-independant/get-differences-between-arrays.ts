/**
 * Compares two arrays of objects and returns only the entries that are new or have changed.
 * New objects (not present in initialArray) are returned in full.
 * Changed objects include only the modified fields plus the identifier field.
 *
 * @param initialArray The original array of objects to compare against.
 * @param newArray The updated array of objects to compare.
 * @param idColName The property name used as a unique identifier for matching objects between arrays.
 * @returns Array of objects representing additions and field-level changes since initialArray.
 */
function main(
  workbook: ExcelScript.Workbook,
  initialArray: Record<string, string>[],
  newArray: Record<string, string>[],
  idColName: string
): {}[] {
  const output: { [key: string]: string }[] = [];

  for (const newObj of newArray) {
    const id: string = newObj[idColName];
    const initialObj = initialArray.find(o => o[idColName] === id);

    // No match found: object is brand new so include in full
    if (!initialObj) {
      output.push({ ...newObj });
      continue;
    }

    // Compare properties and collect only those that have changed
    const diff: { [key: string]: string } = {};

    for (const key of Object.keys(newObj)) {
      if (key === idColName) continue;

      // id is always added if there are changes
      const newVal: string = newObj[key];
      const oldVal: string = initialObj[key];
      if (newVal !== oldVal) {
        diff[key] = newVal
      }
    }

    if (Object.keys(diff).length > 0) {
      diff[idColName] = id
      output.push({ ...diff });
    }
  }

  return output;
}