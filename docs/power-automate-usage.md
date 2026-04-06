# Power Automate Usage

This repo is intended for use with Excel Online (Business) Run script actions.

## Prerequisites

- Office Scripts enabled in the tenant.
- Workbook stored in OneDrive or SharePoint.
- Tables/sheets referenced by script parameters already exist where required.

## Basic Flow Pattern

1. Prepare input values from trigger or prior actions.
2. (Optional) Use Compose actions to build JSON strings.
3. Add Run script and select the workbook + script.
4. Map script parameters exactly.
5. Read result from `outputs('Run_script')?['body/result']`.

## JSON Parameter Tips

- Pass JSON as valid string content for parameters like `inputJson` and `updatesJson`.
- Keep property names aligned with table column names for table scripts.
- Prefer ISO dates (YYYY-MM-DD) for date-like text fields.

## Troubleshooting

- `Table ... not found`: confirm table names match exactly.
- `Column ... not found`: verify source schema and spelling.
- JSON parse errors: validate payload with a Compose action before Run script.
- Empty/no-op result: confirm filters, key values, and worksheet context.
