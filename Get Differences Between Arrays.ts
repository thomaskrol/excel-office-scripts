function main(
  workbook: ExcelScript.Workbook,
  initialArray: Array<object>,
  newArray: Array<object>,
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