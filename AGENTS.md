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
- **Preserves Excel Formatting**: Keeps fonts, colors, borders, merges, and formulas from the original template when expanding rows.
- **Conditional Rendering**: Automatically removes rows if the corresponding array is empty.
- **Safe Missing Key Handling**: Leaves markers untouched if a data key is missing, allowing for easy debugging.
- **Type-Safe API**: Written in TypeScript with full type definitions.
- **ESM/CJS Support**: Ships in both modern ESM and CommonJS formats.

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
- Supports nested paths: `{{user.profile.email}}`.
- If the key is missing, the marker remains in the cell as-is (e.g., `{{missing.key}}`).
- If the value is `null` or `undefined`, the cell becomes empty (`""`).

### 2. Array-Based Row Repetition: `[[array.key]]`

- Used inside a row to indicate that the entire row should be repeated for each item in the `array`.
- The array name (`array`) must exist at the root of the data object and must be an array.
- Each occurrence of `[[array.key]]` in the row is replaced with the value of `item.key` from the current iteration.
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
