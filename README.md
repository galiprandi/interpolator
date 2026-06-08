# @interpolator

A family of lightweight, format-specific template interpolators for structured data. Fill your documents using a simple, declarative syntax while preserving all the original formatting, formulas, and styles.

## Available Packages

- **[@interpolator/xlsx](./packages/xlsx)** – Fill Excel templates (`.xlsx`) with `{{}}` and `[[]]` markers.
- **[@interpolator/core](./packages/core)** – Shared logic for marker parsing and path resolution.

---

## Installation

```bash
pnpm add @interpolator/xlsx
# or
npm install @interpolator/xlsx
```

---

## Guide: How to use `@interpolator`

This guide focuses on `@interpolator/xlsx`, the primary tool for Excel template interpolation.

### 1. Basic Concepts

The library uses two types of markers to tell the engine where to put your data:

#### Single Value Interpolation: `{{path.to.key}}`
Used for static values that appear once in your document (e.g., a customer name, a date, or a total).
- **Nested Paths**: Supports dot notation like `{{user.profile.name}}`.
- **Missing Data**: If a key is missing, the marker `{{key}}` stays in the cell for easier debugging.
- **Null Values**: If a value is `null` or `undefined`, the cell is cleared.
- **Transforms**: Use the pipe operator `|` to transform values: `{{name | upper}}`. Supported: `upper`, `lower`, `capitalize`, `trim`, `trimStart`, `trimEnd`, `camelCase`, `pascalCase`, `snakeCase`, `kebabCase`, `titleCase`, `initials`, `json`, `join`, `unique`, `first`, `last`, `length`, `plural`, `round`, `floor`, `ceil`, `abs`, `reverse`, `sort`, `compact`, `sum`, `avg`, `min`, `max`, `empty`, `notempty`, `boolean`, `keys`, `values`, `lines`, `flat`. You can chain them: `{{title | trim | capitalize}}`.

#### Array-Based Row Expansion: `[[array.key]]`
Used to repeat an entire row for every item in an array.
- **Dynamic Growth**: The worksheet automatically grows as rows are inserted.
- **Item Context**: Inside an expansion row, `[[items.name]]` refers to the `name` property of the current item.
- **Primitive Arrays**: Use `[[items]]` if your array contains strings or numbers directly.
- **Empty Arrays**: If the array is empty, the entire template row is removed.

---

### 2. Powerful Features

#### Smart Type Preservation
Unlike other libraries that convert everything to strings, `@interpolator` is smart. If a cell contains **only** a marker (e.g., `{{amount}}`), the output cell will retain the original data type (Number, Date, Boolean).
- Input: `{ amount: 1250.50 }` -> Output: `1250.50` (Excel Number type).

#### Automatic Formula Adjustment
If your template row has formulas that reference other cells in the same row (e.g., `=B2*C2`), `@interpolator` will automatically adjust them for every new row created (e.g., `=B3*C3`, `=B4*C4`).

#### Dynamic Worksheet Names
You can use `{{key}}` markers directly in the names of your Excel tabs.
- Template Sheet Name: `Report {{month}}`
- Data: `{ month: 'January' }`
- Result: `Report January`

#### Style & Merge Preservation
Fonts, colors, borders, and even horizontal merged cells are copied from the template row to every newly generated row.

---

### 3. Usage Example

```ts
import { interpolateXlsx } from '@interpolator/xlsx';
import { readFileSync, writeFileSync } from 'fs';

const template = readFileSync('./template.xlsx');
const data = {
  client: 'John Doe',
  items: [
    { desc: 'Widget A', qty: 2, price: 10 },
    { desc: 'Widget B', qty: 1, price: 50 },
  ]
};

const result = await interpolateXlsx({ template, data });
writeFileSync('./output.xlsx', result);
```

---

## Use Cases

### Case 1: Automated Invoicing
Create a beautiful Excel invoice with your company branding. Use `{{client}}` for header info and `[[items.description]]`, `[[items.amount]]` for the line items. The total formula at the bottom will still work perfectly as the table expands.

### Case 2: Multi-Sheet Reports
Generate a single workbook where each sheet is named after a specific region or department using `{{department.name}}` in the tab title.

### Case 3: Data Exports with Counters
Need a numbered list or cell references? Use the special metadata markers:
- `[[items.$index]]`: 0, 1, 2...
- `[[items.$number]]` or `[[items.$index1]]`: 1, 2, 3...
- `{{$cell}}` or `[[items.$cell]]`: A1, B2, etc. (Current cell address)
- `{{$colLetter}}` or `[[items.$colLetter]]`: A, B, AA, etc. (Current column letter)

Example cell: `[[items.$number]]. [[items.name]]` produces "1. My Item".

---

## Error Handling

- **Mixed Array Keys**: If you try to use `[[items.a]]` and `[[other.b]]` in the same row, the library will throw a descriptive error.
- **Type Safety**: If a marker expects an array but receives an object, you'll get a clear error message indicating exactly which worksheet and row caused the issue.

---

## License
MIT
