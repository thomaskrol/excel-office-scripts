# Contributing

## Adding a new script

### Folder placement

- `scripts/tables/`: operations that target a named table
- `scripts/worksheets/`: operations that target a sheet or the workbook
- `scripts/workbook-independant/`: pure utility logic that does not interact with a workbook

### File naming

- kebab-case, action-oriented: `verb-noun.ts` (e.g. `update-a-row.ts`)
- Match the README heading in title case

### Script skeleton

```typescript
/**
 * One-line description of what the script does.
 *
 * @param requiredParam Description of requiredParam.
 * @param optionalParam Description of optionalParam.
 * @returns Description of the return value (omit if void).
 */
function main(
  workbook: ExcelScript.Workbook,
  requiredParam: string,
  optionalParam?: number
): { message: string } | void {
  // validate inputs
  // perform work
  // return result (structured objects preferred for Power Automate compatibility)
}
```

### Conventions

- Required parameters before optional ones
- Use explicit union types for constrained options: `"New sheet" | "Existing sheet"`
- Throw `Error` with a clear, actionable message rather than returning error strings
- Prefer returning `{ message, ... }` objects over plain strings so flow expressions stay consistent
- Clear active filters before operating on filtered tables to avoid row-visibility issues

### README entry

After adding a script, add a section to [README.md](README.md) under the correct alphabetical position:

```markdown
### [Script Name](scripts/<folder>/script-name.ts)

One-line description.

| Parameter | Type   | Description  |
| --------- | ------ | ------------ |
| `param`   | string | What it does |

Example input:

\`\`\`json
{ "param": "value" }
\`\`\`

Example output:

\`\`\`json
{ "message": "..." }
\`\`\`

---
```
