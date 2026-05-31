# Project Agents Context

This document provides high-level context for agents and contributors working on the `@interpolator/xlsx` package within the `interpolator` monorepo (GitHub repository: `interpolator`).

---

# `@interpolator/xlsx` – Template Interpolation for Excel Files

## Overview

`@interpolator/xlsx` is a lightweight library (with `exceljs` as its only runtime dependency) designed to **fill Excel templates with structured data**, enabling the generation of dynamic, formatted spreadsheets from JSON inputs.

It provides a simple, declarative syntax for:

- Single-value interpolation.
- Repeating rows based on arrays.
- Preserving formatting, formulas, and styles from the original template.

It is part of the `@interpolator` family of libraries, each targeting a specific document format (e.g., XLSX, Markdown, DOCX).

Agents should assume:

- The main responsibility of this package is to take a template XLSX (as a `Buffer`) and a plain JS object, and return a new XLSX buffer with markers resolved according to the rules below.
- Non-specified behavior should be conservative (e.g., when in doubt, keep original content/markers intact rather than guessing values).

---

## Key Features

- **Declarative Template Syntax**: Use `{{key}}` for single-value interpolation and `[[array.key]]` for repeating rows based on arrays.
- **Smart Type Preservation**: Automatically preserves original data types (Number, Date, Boolean) when a cell contains only a single marker.
- **Worksheet Name Interpolation**: Use `{{key}}` in worksheet names to generate dynamic sheet titles.
- **Preserves Excel Formatting**: Keeps fonts, colors, borders, merges, and formulas from the original template when expanding rows.
- **Conditional Rendering**: Automatically removes rows if the corresponding array is empty.
- **Safe Missing Key Handling**: Leaves markers untouched if a data key is missing, allowing for easy debugging.
- **Type-Safe API**: Written in TypeScript with full type definitions.
- **ESM/CJS Support**: Ships in both modern ESM and CommonJS formats.

---

## Roadmap & Future Directions

### `@interpolator/markdown` (Next Priority)
We are planning to extend the ecosystem with a Markdown interpolator. Key features will include:
- Table expansion from arrays.
- Conditional sections.
- Nested list interpolation.
- Preservation of Markdown structure and metadata.

---

## Installation

```bash
pnpm add @interpolator/xlsx
```

---

## Usage

### Basic Example

```ts
import { interpolateXlsx } from '@interpolator/xlsx';
import { readFileSync, writeFileSync } from 'fs';

// Load your Excel template
const templateBuffer = readFileSync('./invoice-template.xlsx');

// Define your data
const data = {
  client: { name: 'Germán Aliprandi', email: 'g@example.com' },
  items: [
    { id: '001', description: 'Whisky Alberour 12', qty: 2, price: 12000 },
    { id: '002', description: 'Glass set', qty: 1, price: 3000 }
  ],
  total: 27000
};

// Interpolate the template
const resultBuffer = await interpolateXlsx({
  template: templateBuffer,
  data
});

// Write the result to a file
writeFileSync('./filled-invoice.xlsx', resultBuffer);
```

### Template Example

Your Excel template (`.xlsx`) might look like this:

| A                              | B                 | C         | D               |
|--------------------------------|-------------------|-----------|-----------------|
| Invoice for: `{{client.name}}` |                   |           |                 |
| ID                             | Description       | Quantity  | Price           |
| `[[items.id]]`                 | `[[items.description]]` | `[[items.qty]]` | `[[items.price]]` |
|                                |                   |           | `=C2*D2`        |

After interpolation with the data above, the result would be:

| A                              | B                 | C         | D               |
|--------------------------------|-------------------|-----------|-----------------|
| Invoice for: Germán Aliprandi  |                   |           |                 |
| ID                             | Description       | Quantity  | Price           |
| 001                            | Whisky Alberour 12| 2         | 12000           |
| 002                            | Glass set         | 1         | 3000            |
|                                |                   |           | `=C3*D3`        |

---

## Syntax Guide

### 1. Single Value Interpolation: `{{key}}`

- Used for static values that do not change per row.
- **Type Preservation**: If the cell contains *only* the `{{key}}` marker, the resulting cell will have the same type as the data (e.g., Number, Date, Boolean).
- Supports nested paths: `{{user.profile.email}}`.
- **Default values**: `{{path || fallback}}`. Fallback can be a literal string or another path.
- **Transforms**: Use the pipe operator `|` to transform values: `{{name | upper}}`. Supported: `upper`, `lower`, `capitalize`, `trim`, `camelCase`. You can chain them: `{{title | trim | capitalize}}`.
- If the key is missing and no default is provided, the marker remains in the cell as-is.
- If the value is `null` or `undefined` and no default is provided, the cell becomes empty (`""`).

### 2. Array-Based Row Repetition: `[[array.key]]` or `[[array]]`

- Used inside a row to indicate that the entire row should be repeated for each item in the `array`.
- **Conditional Rows**: If the resolved value of the array key is a **boolean**, the row is rendered once (if `true`) or removed (if `false`).
- **Type Preservation**: If a cell contains *only* a `[[ ]]` marker, it will preserve the original data type (e.g., Number, Date, Boolean).
- The array name (`array`) must exist at the root of the data object and must be an array.
- Each occurrence of `[[array.key]]` in the row is replaced with the value of `item.key` from the current iteration.
- `[[array]]` (without property) is replaced by the item itself (useful for primitive arrays).
- Special metadata markers are supported:
  - `[[array.$index]]`: 0-based index (Number).
  - `[[array.$index1]]` or `[[array.$number]]`: 1-based index (Number).
  - `[[array.$first]]` or `[[array.$isFirst]]`: `true` for the first item.
  - `[[array.$last]]` or `[[array.$isLast]]`: `true` for the last item.
  - `[[array.$length]]`: Total number of items in the array.
  - `[[array.$even]]` or `[[array.$isEven]]`: `true` for even-numbered rows.
  - `[[array.$odd]]` or `[[array.$isOdd]]`: `true` for odd-numbered rows.
  - `[[array.$row]]` or `[[array.$rowNumber]]`: Current row number.
  - `[[array.$rowIndex]]`: Current 0-based row index.
  - `[[array.$col]]` or `[[array.$colNumber]]`: Current column number.
  - `[[array.$colIndex]]`: Current 0-based column index.
  - `[[array.$colLetter]]` or `[[array.$columnLetter]]`: Current column letter.
  - `[[array.$isEvenCol]]`: `true` for even-numbered columns.
  - `[[array.$isOddCol]]`: `true` for odd-numbered columns.
- If the array is empty (`[]`), the row is removed from the output.
- If `array` does not exist in the data, markers are left as-is.
- If an item in the array does not have the specified key, the marker remains (e.g., `[[items.missing]]`).
- If an item's key value is `null` or `undefined`, the cell becomes empty.

### 3. Mixed Contexts in One Row

A row can contain both `{{}}` and `[[ ]]` markers:

```text
| [[items.id]] | {{client.name}} | [[items.qty]] |
```

In this case:

- `[[items.id]]` and `[[items.qty]]` are replaced by values from the current `item`.
- `{{client.name}}` is replaced by the same value from the root data object in every repeated row.

### 4. Built-in Context Markers

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

---

## API

### `interpolateXlsx(options)`

```ts
async function interpolateXlsx(options: {
  template: Buffer;            // XLSX file as Buffer
  data: Record<string, any>;   // Data to interpolate
}): Promise<Buffer>;           // Resulting XLSX file as Buffer
```

- **`template`**: A `Buffer` containing a valid `.xlsx` file.
- **`data`**: A plain JavaScript object containing the data to interpolate.
- **Returns**: A `Promise` resolving to a `Buffer` of the filled XLSX file.

Agents should:

- Preserve this contract when refactoring.
- Avoid introducing breaking changes without explicit versioning.
- Keep `interpolateXlsx` side-effect free (pure function given its inputs, aside from CPU/memory usage).

---

## Behavior Summary

| Scenario                                      | Behavior                               |
|-----------------------------------------------|----------------------------------------|
| `{{user.name}}`, `user.name` exists           | Value is interpolated                  |
| `{{user.name}}`, `user.name` missing          | Cell contains `{{user.name}}`          |
| `{{user.name}}`, `user.name` is `null`        | Cell becomes empty (`""`)             |
| `[[items.id]]`, `items` is `[]`               | Row is removed                         |
| `[[items.id]]`, `items` missing               | Cell contains `[[items.id]]`           |
| `[[items.id]]`, `items` is not an array       | Throws an error                        |
| `[[items.id]]`, `item` has no `id`            | Cell contains `[[items.id]]`           |
| `[[items.id]]`, `item.id` is `null`           | Cell becomes empty (`""`)             |
| Row has formulas                              | Formulas are preserved/adjusted        |
| Row has formatting                            | Formatting is copied to new rows       |

---

## Typical Use Cases

- **Invoices & Reports**: Generate personalized invoices, reports, or statements from templates.
- **Data Export**: Convert JSON data into formatted Excel files with consistent styling.
- **Bulk Mailing Labels**: Create sheets of labels or cards from a list of data.
- **Dynamic Dashboards**: Populate Excel-based dashboards with real-time data.

---

## License

MIT
