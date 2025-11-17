# @interpolator/xlsx

Excel template interpolation for Node.js. Fill `.xlsx` templates using a simple marker syntax and a plain JavaScript object, while preserving styles and formulas.

- Interpolate single values with `{{key}}`.
- Repeat rows for arrays with `[[array.key]]`.
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
- If the key (or any intermediate segment) does not exist, the **marker is left as-is**.
- If the resolved value is `null` or `undefined`, the cell becomes an empty string (`""`).

### `[[]]` – array row expansion

Use `[[array.key]]` in a row to mark it as **repeatable**:

- The array name (e.g. `items` in `[[items.id]]`) must exist at the root of `data` and be an array.
- The row containing `[[items.*]]` is removed and replaced by **N rows**, one per item in `data.items`.
- Each `[[items.prop]]` is resolved against the corresponding item.
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

## Formatting, formulas and merges

- Cell styles (font, fill, border, alignment, etc.) from the template row are copied to each new row.
- Formulas in the template row are preserved and adjusted when they reference the template row number
  (e.g. `=B2*C2` becomes `=B3*C3` in the next cloned row).
- Merged cells: the implementation attempts to replicate merges for cloned rows, but due to current
  ExcelJS limitations, this behavior is **best-effort** and not fully guaranteed across all scenarios.

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
