# Naming Conventions

Use these conventions to keep new scripts consistent and easy to find.

## File Names

- Use kebab-case.
- Keep names action-oriented (verb + object), e.g. `update-a-row.ts`.
- Match script name in README title case.

## Script Signature

- Export a `main` function with `workbook: ExcelScript.Workbook` first.
- Keep required parameters before optional ones.
- Use explicit union types for constrained options.

## Documentation

- Include JSDoc on every script.
- Document all parameters and return shape.
- Mention Power Automate behavior when relevant.

## Return Values

- Prefer structured objects for flow-friendly parsing.
- Include `message` and/or `success` for operational clarity.
- For errors, throw clear messages with actionable context.

## Folder Placement

- `scripts/tables`: Table-focused operations.
- `scripts/worksheets`: Sheet/workbook operations.
- `scripts/workbook-independant`: Utility logic that does not depend on workbook ranges.
