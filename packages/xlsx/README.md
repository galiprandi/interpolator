# @interpolator/xlsx

Excel template interpolation for Node.js. Fill `.xlsx` templates using a simple marker syntax and a plain JavaScript object, while preserving styles and formulas.

- Interpolate single values with `{{key}}`.
- Repeat rows for arrays with `[[array.key]]` (objects) or `[[array]]` (primitives).
- Preserve existing formatting and formulas in the template.
- Keep markers when data is missing for easier debugging.

## Installation

```bash
pnpm add @interpolator/xlsx
# or
npm install @interpolator/xlsx
# or
yarn add @interpolator/xlsx
```

## Basic usage

```ts
import { interpolateXlsx } from '@interpolator/xlsx';
import { readFileSync, writeFileSync } from 'node:fs';

// 1) Load your Excel template as a Buffer
const template = readFileSync('./invoice-template.xlsx');

// 2) Provide your data object
const data = {
  client: { name: 'Germán Aliprandi', email: 'galiprandi@gmail.com' },
  items: [
    { id: '001', description: 'Whisky Alberour 12', qty: 2, price: 12000 },
    { id: '002', description: 'Whisky Johnnie Walker Green', qty: 1, price: 15000 },
  ],
  total: 27000,
};

// 3) Interpolate the template
const resultBuffer = await interpolateXlsx({ template, data });

// 4) Save the result
writeFileSync('./filled-invoice.xlsx', resultBuffer);
```

### Example template

Given a worksheet like:

| A                              | B           | C         | D               |
|--------------------------------|-------------|-----------|-----------------|
| Invoice for: `{{client.name}}` |             |           |                 |
| ID                             | Description | Quantity  | Price           |
| `[[items.id]]`                 | `[[items.description]]` | `[[items.qty]]` | `[[items.price]]` |
|                                |             |           | `=C3*D3`        |

After interpolation, the rows for `items` will be expanded and the formulas preserved/adjusted per row.

## Marker syntax

### `{{}}` – single value interpolation

Use `{{path.to.value}}` for values that do not depend on the current row:

- Resolved against the root `data` object.
- Supports deep paths like `{{user.profile.email}}`.
- **Default values**: Use `||` to provide a fallback if a value is missing or null.
  - `{{user.name || N/A}}` -> renders "N/A" if `user.name` is missing or null.
  - `{{user.city || user.backupCity}}` -> tries to resolve `user.backupCity` if `user.city` is not found.
- If the key (or any intermediate segment) does not exist and no default is provided, the **marker is left as-is**.
- If the resolved value is `null` or `undefined` and no default is provided, the cell becomes an empty string (`""`).

### `[[]]` – array row expansion

Use `[[array.key]]` (for objects) or `[[array]]` (for primitives) in a row to mark it as **repeatable**:

- The array name (e.g. `items` in `[[items.id]]` or `[[categories]]`) must exist at the root of `data` and be an array.
- The row containing `[[items.*]]` or `[[items]]` is removed and replaced by **N rows**, one per item in the array.
- `[[array.prop]]` is resolved against the current item properties.
- `[[array]]` (without property) is resolved to the item itself (useful for arrays of strings or numbers).
- **Conditional rows**: Use a boolean value for the array key to conditionally show or hide a row.
  - `[[showPremiumFeatures]]` -> if `true`, the row is rendered once (with markers removed); if `false`, the row is removed.
- The same row can mix `{{}}` and `[[]]`:

  ```text
  | [[items.id]] | {{client.name}} | [[items.qty]] |
  ```

  - `{{client.name}}` is taken from the root data in every repeated row.
  - `[[items.id]]`, `[[items.qty]]` use the current item.

#### Behavior for missing or invalid arrays

- If the array key **does not exist** in `data`:
  - The row is **not** expanded or removed.
  - Markers like `[[items.id]]` are kept as-is (useful for debugging).
- If the array key exists and is an **empty array (`[]`)**:
  - The row is **removed** from the output (no rows for that section).
- If the array key exists but is **not an array**:
  - `interpolateXlsx` throws an error with context, e.g.:

    ```text
    [[items.*]] requires "items" to be an array in worksheet "Sheet1", row 5. Received: object
    ```

#### Behavior for item properties

- If an item in the array **does not have** the referenced property:
  - The marker (e.g. `[[items.missing]]`) is left as-is.
- If the item property is `null` or `undefined`:
  - The cell becomes an empty string (`""`).

#### Special index markers

You can use these special property paths within an array context to include indices or counters:

- `[[array.$index]]`: The 0-based index of the current item (0, 1, 2, ...).
- `[[array.$index1]]` or `[[array.$number]]`: The 1-based index of the current item (1, 2, 3, ...).
- `[[array.$first]]` or `[[array.$isFirst]]`: Boolean flag (`true`/`false`) for the first item.
- `[[array.$last]]` or `[[array.$isLast]]`: Boolean flag (`true`/`false`) for the last item.
- `[[array.$even]]` or `[[array.$isEven]]`: Boolean flag (`true`/`false`) for even items (1-indexed).
- `[[array.$odd]]` or `[[array.$isOdd]]`: Boolean flag (`true`/`false`) for odd items (1-indexed).
- `[[array.$length]]`: The total number of items in the array.
- `[[array.$cell]]`: The current cell address (e.g. A2, B3).
- `[[array.$colLetter]]` or `[[array.$columnLetter]]`: The current column letter (e.g. A, B, AA).
- `[[array.$row]]` or `[[array.$rowNumber]]`: The current row number (1-indexed).
- `[[array.$rowIndex]]`: The current row index (0-indexed).
- `[[array.$col]]` or `[[array.$colNumber]]`: The current column number (1-indexed).
- `[[array.$colIndex]]`: The current column index (0-indexed).
- `[[array.$isEvenCol]]`: Boolean flag for even columns (1-indexed).
- `[[array.$isOddCol]]`: Boolean flag for odd columns (1-indexed).

Example: `[[items.$number]] of [[items.$length]]: [[items.name]]` will produce "1 of 10: First Item", etc.

### Built-in context markers

These markers can be used both in `{{}}` and `[[]]` contexts:

- `{{$now}}`: Current date and time.
- `{{$year}}`: Current year (e.g. 2024).
- `{{$month}}`: Current month (1-12).
- `{{$day}}`: Current day of the month (1-31).
- `{{$hour}}`: Current hour (0-23).
- `{{$minute}}`: Current minute (0-59).
- `{{$second}}`: Current second (0-59).
- `{{$weekday}}`: Current day of the week (0-6, where 0 is Sunday).
- `{{$isHeader}}`: `true` if the marker is in the first row.
- `{{$sheet}}` or `{{$sheetName}}`: Current worksheet name.
- `{{$sheetIndex}}`: 0-based worksheet index.
- `{{$sheetNumber}}`: 1-based worksheet index.
- `{{$totalSheets}}`: Total number of sheets in the workbook.
- `{{$isFirstSheet}}`: `true` for the first sheet.
- `{{$isLastSheet}}`: `true` for the last sheet.
- `{{$row}}` or `{{$rowNumber}}`: Current row number (1-indexed).
- `{{$rowIndex}}`: Current row index (0-indexed).
- `{{$col}}` or `{{$colNumber}}`: Current column number (1-indexed).
- `{{$colIndex}}`: Current column index (0-indexed).
- `{{$isEven}}`, `{{$even}}` or `{{$isEvenRow}}`: `true` for even-numbered rows.
- `{{$isOdd}}`, `{{$odd}}` or `{{$isOddRow}}`: `true` for odd-numbered rows.
- `{{$isEvenCol}}`: `true` for even-numbered columns.
- `{{$isOddCol}}`: `true` for odd-numbered columns.
- `{{$colLetter}}` or `{{$columnLetter}}`: Current column letter (e.g. A, B, Z, AA).
- `{{$cell}}`: Current cell address (e.g. A1, B2).

## Formatting, formulas and merges

- Cell styles (font, fill, border, alignment, etc.) from the template row are copied to each new row.
- Formulas in the template row are preserved and adjusted when they reference the template row number
  (e.g. `=B2*C2` becomes `=B3*C3` in the next cloned row).
- Merged cells:
  - Merges that exist in template rows **without** `[[]]` markers are preserved as-is after interpolation.
  - When repeating rows that contain `[[]]`, the implementation attempts to replicate **simple horizontal**
    merges from the template row into each cloned row, but this is **best-effort** and relies on current
    ExcelJS behavior.
  - Vertical merges (spanning multiple rows) or merges that cross repeated and non-repeated rows are
    **not supported** and may behave unpredictably.

## Error behavior

- Missing root keys for `{{}}`: markers are preserved (no error).
- Missing array keys for `[[]]`: markers are preserved and rows are not repeated or removed.
- Empty arrays: rows marked with `[[]]` are removed.
- Non-array values used with `[[]]`: an error is thrown with worksheet and row context.
- Future versions may also validate mixed array keys in the same row (e.g. `[[items.*]]` and `[[payments.*]]`).

## API

```ts
import type { Buffer } from 'node:buffer';

interface InterpolateXlsxOptions {
  template: Buffer;
  data: Record<string, any>;
}

declare function interpolateXlsx(options: InterpolateXlsxOptions): Promise<Buffer>;
```

- **template**: a `Buffer` containing a valid `.xlsx` file.
- **data**: plain JavaScript object with the data to interpolate.
- Returns a **Promise** that resolves to a `Buffer` with the interpolated workbook.

## Project status

This package is part of the `interpolator` monorepo and is under active development.
Behavior is driven by the acceptance criteria in `ACCEPTANCE.md` and the roadmap in
`packages/xlsx/ROADMAP.md`. Expect the public API to remain stable while implementation
and coverage improve over time.
