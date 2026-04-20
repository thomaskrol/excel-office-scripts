/**
  * Highlights table column headers based on column name matching.
  * 
  * @param tableName Name of the table to process
  * @param highlightColor Header fill colour in hex (#RRGGBB or RRGGBB) or named HTML colour (e.g., "orange");
  * @param matchType Matching strategy: "List of column names" for exact matches or "RegEx" for pattern matching
  * @param columnNamesArray Array of exact column names to match (required when matchType is "List of column names");
  * @param regexPattern Regular expression pattern to match column names (required when matchType is "RegEx");
  * @param regexFlags Optional regex flags (g, i, m, s, u, y, d);
  */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  highlightColor: string,
  matchType: "List of column names" | "RegEx" = "List of column names",
  columnNamesArray?: Array<string>,
  regexPattern?: string,
  regexFlags?: string
) {
  // make sure required optional parameters are present
  switch (matchType) {
    case "List of column names": {
      if (!columnNamesArray) {
        throw new Error("Parameter columnNamesArray is required when matchType is 'List of column names'.");
      }
      break;
    }
    case "RegEx": {
      if (!regexPattern) {
        throw new Error("Parameter regexPattern is required when matchType is 'RegEx'.");
      }
      break;
    }
    default: {
      throw new Error("Parameter matchType has an unrecognised value. Valid values are 'List of column names' and 'RegEx'.");
    }
  }

  // validate regexFlags
  if (matchType === "RegEx" && regexFlags) {
    const validFlags = ['g', 'i', 'm', 's', 'u', 'y', 'd'];
    const invalidFlags = regexFlags.split("").filter((f) => !validFlags.includes(f));

    if (invalidFlags.length > 0) {
      throw new Error(`Invalid regex flags: ${invalidFlags.join(', ')}. Valid flags are: ${validFlags.join(', ')}`);
    }
  }

  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table '${tableName}' not found.`);
  }

  const tableCols = table.getColumns();
  // use .includes() or regex to determine which columns to highlight
  let colsToHighlight: ExcelScript.TableColumn[] = [];
  tableCols.forEach((col) => {
    const name = col.getName();
    const isMatch = matchType === "List of column names" ? columnNamesArray.includes(name) : new RegExp(regexPattern, regexFlags || undefined).test(name);

    if (isMatch) {
      colsToHighlight.push(col);
    }
  });

  // does not work when called from Power Automate despite working in Excel
  // const colsToHighlight = tableCols.filter((col) => {
  // const name = col.getName();
  // new RegExp(regexPattern, regexFlags || undefined).test(name);
  // });

  // highlight header row of identified columns
  colsToHighlight.forEach((col) => {
    col.getHeaderRowRange().getFormat().getFill().setColor(highlightColor);
  });

  // check if any columns were not found
  if (matchType === "List of column names") {
    const highlightedColNames = colsToHighlight.map((col) => col.getName());
    const missingCols = columnNamesArray.filter((name) => {
      return !highlightedColNames.includes(name)
    });

    return {
      "message": "Successfully highlighted matched columns.",
      "notFoundColumns": missingCols || []
    }
  }

  return {
    "message": "Successfully highlighted matched columns."
  }
}