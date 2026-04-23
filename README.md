# excel-office-scripts

A collection of [Office Scripts](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel) for Microsoft Excel, designed to be called from [Power Automate](https://learn.microsoft.com/en-us/office/dev/scripts/develop/power-automate-integration) flows to extend automation beyond what standard actions provide.

## Usage

Each script is used as a **Run script** or **Run script from SharePoint library** action in the Excel Online (Business) connector. Prerequisites:

- Office Scripts enabled in the tenant
- Workbook stored in OneDrive or SharePoint
- Tables/sheets referenced by script parameters already exist

Script outputs can be read in later flow actions using `outputs('Run_script')?['body/result']`.

### JSON parameter tips

- Parameters like `inputJson` and `updatesJson` must be passed as a JSON string, not an object. Use a Compose action to build the string first.
- Keep property names aligned exactly with their respective table/column names.
- Prefer ISO dates (`YYYY-MM-DD`) for date-like text fields.

### Troubleshooting

| Symptom | Likely cause |
| --- | --- |
| `Table "..." not found` | Table name mismatch, check exact casing and spacing |
| `Column "..." not found` | Column name mismatch or table schema has changed |
| JSON parse error | Malformed string passed to a JSON parameter, validate with a Compose action first |
| Script runs but nothing changes | Active filter hiding rows, wrong key value, or wrong sheet context |

---

## Repository Layout

Scripts are grouped by category, and each script lives in its own subfolder containing the source `.ts` file and the generated `.osts` file:

```
scripts/
  tables/
    add-rows-to-table/
      add-rows-to-table.ts
      add-rows-to-table.osts
    ...
  worksheets/
  workbook-independant/
```

- `scripts/tables/`: table-focused Office Scripts
- `scripts/worksheets/`: worksheet/workbook scripts
- `scripts/workbook-independant/`: utility scripts that do not require table/sheet context

The `.osts` files are generated automatically by a GitHub Actions workflow on every push to `main`. Do not edit them by hand.

See [CONTRIBUTING.md](CONTRIBUTING.md) for conventions when adding new scripts.

---

## Scripts

### Tables

<details>
<summary>Scripts</summary>

#### [Add Rows To Table](scripts/tables/add-rows-to-table/add-rows-to-table.ts)

Adds one or more rows to an existing table from a JSON array.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table name |
| `inputJson` | string | JSON array whose keys match table columns |

Example input:

```json
{
  "tableName": "Orders",
  "inputJson": "[{\"OrderId\":\"SO-101\",\"Status\":\"New\"}]"
}
```

Example output:

```json
"Rows added successfully."
```

---

#### [Auto Fit Column Widths](scripts/tables/auto-fit-column-widths/auto-fit-column-widths.ts)

Auto-fits all column widths in a table.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table name |

Example input:

```json
{
  "tableName": "Orders"
}
```

---

#### [Clear Cell Contents](scripts/tables/clear-cell-contents/clear-cell-contents.ts)

Clears one cell in a table by data-row index and column name.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table name |
| `columnName` | string | Target column name |
| `rowIndex` | number | Zero-based data row index |

Example input:

```json
{
  "tableName": "Orders",
  "columnName": "Status",
  "rowIndex": 2
}
```

---

#### [Convert Name Cases](scripts/tables/convert-name-cases/convert-name-cases.ts)

Converts all-uppercase/all-lowercase names to Proper Case in selected columns.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table name |
| `columnsToFix` | string[] | Columns to normalize |

Example input:

```json
{
  "tableName": "Contacts",
  "columnsToFix": ["FirstName", "LastName"]
}
```

---

#### [Convert Table Column from Formula to Values](scripts/tables/convert-table-column-from-formula-to-values/convert-table-column-from-formula-to-values.ts)

Converts formulas in a table column to static values.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table name |
| `columnName` | string | Column to convert |
| `fullColumn` | boolean | Convert all rows when true |
| `numberOfRowsFromEnd` | number (optional) | Required when `fullColumn` is false |

Example input:

```json
{
  "tableName": "Orders",
  "columnName": "Total",
  "fullColumn": false,
  "numberOfRowsFromEnd": 50
}
```

---

#### [Create Pivot Table](scripts/tables/create-pivot-table/create-pivot-table.ts)

Creates a pivot table from a source table.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Source table |
| `location` | "New sheet" \| "Existing sheet" | Placement |
| `rowsColumn` | string | Row grouping column |
| `valuesColumns` | string[] | Value columns |
| `valuesOperation` | "Sum" \| "Count" \| "Average" \| "Product" \| "Max" \| "Min" | Aggregation |
| `columnsColumn` | string (optional) | Column grouping |
| `sheetName` | string (optional) | Target sheet |
| `pivotTableName` | string (optional) | Pivot table name |

Example input:

```json
{
  "tableName": "Orders",
  "location": "New sheet",
  "rowsColumn": "Region",
  "valuesColumns": ["Amount"],
  "valuesOperation": "Sum",
  "pivotTableName": "OrdersByRegion"
}
```

Example output:

```json
{
  "message": "Successfully created a pivot table.",
  "createdPivotTableName": "OrdersByRegion",
  "usedSheetName": "Sheet2"
}
```

---

#### [Create Table From JSON](scripts/tables/create-table-from-json/create-table-from-json.ts)

Creates a new table from a JSON array on a chosen sheet/cell.

| Parameter | Type | Description |
| --- | --- | --- |
| `sheetName` | string | Destination worksheet |
| `startCell` | string | Top-left table cell (for example A1) |
| `inputJson` | string | JSON array of objects |
| `tableName` | string (optional) | Name for created table |

Example input:

```json
{
  "sheetName": "Import",
  "startCell": "A1",
  "inputJson": "[{\"OrderId\":\"SO-1\",\"Amount\":120}]",
  "tableName": "ImportedOrders"
}
```

Example output:

```json
{
  "message": "Successfully created table.",
  "createdTableName": "ImportedOrders"
}
```

---

#### [Highlight Specific Table Columns](scripts/tables/highlight-specific-table-columns/highlight-specific-table-columns.ts)

Highlights table headers matched by name list or regex.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table |
| `highlightColor` | string | Fill color |
| `matchType` | "List of column names" \| "RegEx" | Matching mode |
| `columnNamesArray` | string[] (optional) | Required for list mode |
| `regexPattern` | string (optional) | Required for regex mode |
| `regexFlags` | string (optional) | Regex flags |

Example input:

```json
{
  "tableName": "Orders",
  "highlightColor": "#FFC000",
  "matchType": "List of column names",
  "columnNamesArray": ["Status", "Priority"]
}
```

Example output:

```json
{
  "message": "Successfully highlighted matched columns.",
  "notFoundColumns": []
}
```

---

#### [Set Table Rows Height](scripts/tables/set-table-rows-height/set-table-rows-height.ts)

Sets row height for all rows in a table.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table |
| `rowHeight` | number (optional) | Height in points |

Example input:

```json
{
  "tableName": "Orders",
  "rowHeight": 18
}
```

---

#### [Sort Table By Column Name](scripts/tables/sort-table-by-column-name/sort-table-by-column-name.ts)

Sorts a table ascending by a selected column.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table |
| `columnName` | string | Sort key column |

Example input:

```json
{
  "tableName": "Orders",
  "columnName": "OrderDate"
}
```

---

#### [Update a Row](scripts/tables/update-a-row/update-a-row.ts)

Updates one row identified by a key column value.

| Parameter | Type | Description |
| --- | --- | --- |
| `tableName` | string | Target table |
| `keyColumnName` | string | Lookup column |
| `keyValue` | string | Lookup value |
| `updatesJson` | string | JSON object of column:value updates |

Example input:

```json
{
  "tableName": "Orders",
  "keyColumnName": "OrderId",
  "keyValue": "SO-101",
  "updatesJson": "{\"Status\":\"Done\",\"Owner\":\"Ops\"}"
}
```

Example output:

```json
{
  "success": true,
  "message": "Row updated successfully",
  "row": 14
}
```

</details>

---

### Worksheets

<details>
<summary>Scripts</summary>

#### [Delete Sheet](scripts/worksheets/delete-sheet/delete-sheet.ts)

Deletes a worksheet.

| Parameter | Type | Description |
| --- | --- | --- |
| `sheetName` | string | Sheet to delete |

Example input:

```json
{
  "sheetName": "OldData"
}
```

---

#### [Hide Sheet](scripts/worksheets/hide-sheet/hide-sheet.ts)

Hides a worksheet.

| Parameter | Type | Description |
| --- | --- | --- |
| `sheetName` | string | Sheet to hide |

Example input:

```json
{
  "sheetName": "RawData"
}
```

---

#### [Replace in Sheet](scripts/worksheets/replace-in-sheet/replace-in-sheet.ts)

Replaces all matching values in a worksheet.

| Parameter | Type | Description |
| --- | --- | --- |
| `sheetName` | string | Target sheet |
| `oldValue` | string | Value to find |
| `newValue` | string | Replacement value |
| `matchCase` | boolean | Case-sensitive when true |
| `matchEntireCellContents` | boolean | Whole-cell match only when true |

Example input:

```json
{
  "sheetName": "Orders",
  "oldValue": "Pending",
  "newValue": "In Progress",
  "matchCase": false,
  "matchEntireCellContents": true
}
```

</details>

---

### Workbook-independant

<details>
<summary>Scripts</summary>

#### [Get Differences Between Arrays](scripts/workbook-independant/get-differences-between-arrays/get-differences-between-arrays.ts)

Returns objects that are new or changed between two arrays.

| Parameter | Type | Description |
| --- | --- | --- |
| `initialArray` | object[] | Baseline array |
| `newArray` | object[] | Updated array |
| `idColName` | string | Identity key name |

Example input:

```json
{
  "initialArray": [{"id":"1","status":"New"}],
  "newArray": [{"id":"1","status":"Done"},{"id":"2","status":"New"}],
  "idColName": "id"
}
```

Example output:

```json
[
  {"status":"Done","id":"1"},
  {"id":"2","status":"New"}
]
```

---

#### [Regex Operations](scripts/workbook-independant/regex-operations/regex-operations.ts)

Runs regex match/test/replace/group operations on a string.

| Parameter | Type | Description |
| --- | --- | --- |
| `operation` | "all matches" \| "test match" \| "replace" \| "groups" | Operation |
| `searchString` | string | Input string |
| `regexPattern` | string | Regex pattern |
| `regexFlags` | string (optional) | Regex flags |
| `replaceString` | string (optional) | Replacement text for replace |

Example input:

```json
{
  "operation": "replace",
  "searchString": "Order SO-123",
  "regexPattern": "SO-(\\d+)",
  "replaceString": "ID-$1"
}
```

Example output:

```json
"Order ID-123"
```

</details>
