# excel-office-scripts

A collection of [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) for Microsoft Excel, designed to be called from [Power Automate](https://learn.microsoft.com/en-us/office/dev/scripts/develop/power-automate-integration) flows to extend automation beyond what standard actions provide.

## Usage

Each script is used as a **Run script** or **Run script from SharePoint library** action in a Power Automate flow via the **Excel Online (Business)** connector. Select the script by name and map flow variables to its parameters.

Scripts return structured values (JSON objects or primitives) where noted, which can be parsed in subsequent flow actions using `outputs('Run_script')?['body/result']`.

---

## Scripts

### Add Rows To Table

Adds one or more rows to an existing Excel table from a JSON array. Columns not present in the input will be left empty.

| Parameter   | Type   | Description                                                 |
| ----------- | ------ | ----------------------------------------------------------- |
| `tableName` | string | Name of the target table                                    |
| `inputJson` | string | JSON array of objects with keys matching table column names |

**Returns:** `"Rows added successfully."` or an error message string.

---

### Auto Fit Column Widths

Auto-fits all column widths in a table to best fit their contents.

| Parameter   | Type   | Description              |
| ----------- | ------ | ------------------------ |
| `tableName` | string | Name of the target table |

---

### Automatic Calculations

Sets the workbook calculation mode to **automatic**, forcing recalculation of all formulas. Useful when a previous step has set calculation to manual for performance.

_No parameters beyond the workbook._

---

### Clear Cell Contents

Clears the contents of a specific cell in a table identified by column name and row index.

| Parameter    | Type   | Description                                                   |
| ------------ | ------ | ------------------------------------------------------------- |
| `tableName`  | string | Name of the target table                                      |
| `columnName` | string | Name of the column                                            |
| `rowIndex`   | number | Zero-based row index within the table data (excluding header) |

---

### Convert Name Cases

Converts values in specified columns to Proper Case when they are entirely uppercase or entirely lowercase. Handles accented characters common in European names.

| Parameter      | Type     | Description                                   |
| -------------- | -------- | --------------------------------------------- |
| `tableName`    | string   | Name of the target table                      |
| `columnsToFix` | string[] | Column names whose values should be converted |

---

### Convert Table Column from Formula to Values

Replaces formulas in a table column with their calculated static values. Clears active filters before converting to ensure all rows are affected.

| Parameter             | Type               | Description                                                                    |
| --------------------- | ------------------ | ------------------------------------------------------------------------------ |
| `tableName`           | string             | Name of the target table                                                       |
| `columnName`          | string             | Name of the column to convert                                                  |
| `fullColumn`          | boolean            | `true` to convert the entire column; `false` to convert only the last n rows   |
| `numberOfRowsFromEnd` | number_(optional)_ | Number of rows from the end to convert (required when `fullColumn` is `false`) |

---

### Create Pivot Table

Creates a pivot table from an existing table with configurable row hierarchies, value aggregations, and optional column hierarchy.

| Parameter         | Type                                                                     | Description                                                           |
| ----------------- | ------------------------------------------------------------------------ | --------------------------------------------------------------------- |
| `tableName`       | string                                                                   | Name of the source table                                              |
| `location`        | `"New sheet"` \| `"Existing sheet"`                                      | Where to place the pivot table                                        |
| `rowsColumn`      | string                                                                   | Column to use for row grouping                                        |
| `valuesColumns`   | string[]                                                                 | Columns to aggregate                                                  |
| `valuesOperation` | `"Sum"` \| `"Count"` \| `"Average"` \| `"Product"` \| `"Max"` \| `"Min"` | Aggregation function                                                  |
| `columnsColumn`   | string_(optional)_                                                       | Column to use for column grouping                                     |
| `sheetName`       | string_(optional)_                                                       | Target sheet name (or new sheet name when location is `"New sheet"`)  |
| `pivotTableName`  | string_(optional)_                                                       | Name for the pivot table (auto-generated if omitted or already taken) |

**Returns:** `{ message, createdPivotTableName, usedSheetName }`

---

### Create Table From JSON

Creates a new Excel table on a specified worksheet from a JSON array, setting column headers from the JSON keys.

| Parameter   | Type               | Description                                      |
| ----------- | ------------------ | ------------------------------------------------ |
| `sheetName` | string             | Name of the target worksheet                     |
| `startCell` | string             | Top-left cell address for the table (e.g.`"A1"`) |
| `inputJson` | string             | JSON array of objects with consistent keys       |
| `tableName` | string_(optional)_ | Name for the created table                       |

**Returns:** `{ message, createdTableName }`

---

### Delete Sheet

Permanently deletes a worksheet from the workbook.

| Parameter   | Type   | Description                     |
| ----------- | ------ | ------------------------------- |
| `sheetName` | string | Name of the worksheet to delete |

---

### Get Differences Between Arrays

Compares two arrays of objects and returns only entries that are new or have changed. New objects are returned in full; changed objects include only the modified fields plus the identifier.

| Parameter      | Type     | Description                                 |
| -------------- | -------- | ------------------------------------------- |
| `initialArray` | object[] | The original array to compare against       |
| `newArray`     | object[] | The updated array to compare                |
| `idColName`    | string   | Property name used as the unique identifier |

**Returns:** Array of objects representing additions and field-level changes.

---

### Hide Sheet

Hides a worksheet in the workbook.

| Parameter   | Type   | Description                   |
| ----------- | ------ | ----------------------------- |
| `sheetName` | string | Name of the worksheet to hide |

---

### Highlight Specific Table Columns

Highlights table column headers based on column name matching, either from an explicit list or a regular expression.

| Parameter          | Type                                  | Description                                                             |
| ------------------ | ------------------------------------- | ----------------------------------------------------------------------- |
| `tableName`        | string                                | Name of the target table                                                |
| `highlightColor`   | string                                | Fill colour in hex (`#RRGGBB`) or named HTML colour (e.g. `"orange"`)   |
| `matchType`        | `"List of column names"` \| `"RegEx"` | Matching strategy                                                       |
| `columnNamesArray` | string[]_(optional)_                  | Exact column names to highlight (required for `"List of column names"`) |
| `regexPattern`     | string_(optional)_                    | Regex pattern to match column names (required for `"RegEx"`)            |
| `regexFlags`       | string_(optional)_                    | Regex flags (e.g.`"i"` for case-insensitive)                            |

**Returns:** `{ message, notFoundColumns }` — `notFoundColumns` is populated when using `"List of column names"`.

---

### Regex Operations

Performs a regex operation on an input string without needing a spreadsheet cell. Useful for string manipulation within a Power Automate flow.

| Parameter       | Type                                                           | Description                                   |
| --------------- | -------------------------------------------------------------- | --------------------------------------------- |
| `operation`     | `"all matches"` \| `"test match"` \| `"replace"` \| `"groups"` | Operation to perform                          |
| `searchString`  | string                                                         | The string to operate on                      |
| `regexPattern`  | string                                                         | The regex pattern                             |
| `regexFlags`    | string_(optional)_                                             | Regex flags (e.g.`"gi"`)                      |
| `replaceString` | string_(optional)_                                             | Replacement string (required for `"replace"`) |

**Returns:** Matched strings array, boolean, or replaced string depending on operation.

---

### Replace in Sheet

Replaces all occurrences of a value in a worksheet with a new value.

| Parameter                 | Type    | Description                                        |
| ------------------------- | ------- | -------------------------------------------------- |
| `sheetName`               | string  | Name of the target worksheet                       |
| `oldValue`                | string  | Value to search for                                |
| `newValue`                | string  | Replacement value                                  |
| `matchCase`               | boolean | Case-sensitive search (defaults to `false`)        |
| `matchEntireCellContents` | boolean | Match only whole-cell values (defaults to `false`) |

---

### Set Table Rows Height

Sets the row height for all rows in a table (including the header row).

| Parameter   | Type               | Description                            |
| ----------- | ------------------ | -------------------------------------- |
| `tableName` | string             | Name of the target table               |
| `rowHeight` | number_(optional)_ | Height in points (defaults to `14.25`) |

---

### Sort Table By Column Name

Sorts a table in ascending order by the specified column.

| Parameter    | Type   | Description                   |
| ------------ | ------ | ----------------------------- |
| `tableName`  | string | Name of the target table      |
| `columnName` | string | Name of the column to sort by |

---

### Update a Row

Updates one or more column values in a table row identified by a key column value. Returns the Excel row number of the updated row on success.

| Parameter       | Type   | Description                                                                         |
| --------------- | ------ | ----------------------------------------------------------------------------------- |
| `tableName`     | string | Name of the target table                                                            |
| `keyColumnName` | string | Column used to identify the target row                                              |
| `keyValue`      | string | Value in the key column that identifies the row                                     |
| `updatesJson`   | string | JSON object of column-name-to-value mappings (e.g.`{"Status": "Done", "Count": 5}`) |

**Returns:** `{ success, message, row }` — `row` is the 1-based Excel row number of the updated row.
