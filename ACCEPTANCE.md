# Acceptance Criteria: `@interpolator/xlsx`

## 1. Marker syntax

### 1.1. Simple interpolation with `{{}}`

- **Given** a cell with `{{user.name}}`,
- **And** `data` includes `{ user: { name: "Germán" } }`,
- **Then** the cell must contain `"Germán"`.

### 1.2. Interpolation with spaces

- **Given** a cell with `{{ user.name }}`,
- **And** `data` includes `{ user: { name: "Germán" } }`,
- **Then** it must behave the same as `{{user.name}}`.

### 1.3. Deep nested interpolation

- **Given** `{{profile.contact.email}}`,
- **And** data: `{ profile: { contact: { email: "g@a.com" } } }`,
- **Then** the cell must contain `"g@a.com"`.

### 1.4. Root key does not exist

- **Given** `{{user.name}}`,
- **And** `user` does not exist in `data`,
- **Then** the cell must contain the literal string `{{user.name}}`.

### 1.5. Intermediate property does not exist

- **Given** `{{user.profile.email}}`,
- **And** `user` exists but has no `profile` property,
- **Then** the cell must contain `{{user.profile.email}}`.

### 1.6. Value is `null` or `undefined`

- **Given** `{{user.name}}`,
- **And** `user.name` is `null` or `undefined`,
- **Then** the cell must be empty (`""`).

---

## 2. Array interpolation with `[[]]`

### 2.1. Valid non-empty array

- **Given** a row with `[[payments.id]]` and `[[payments.date]]`,
- **And** `payments` is: `[{ "id": "P1", "date": "2025-01-01" }, { "id": "P2", "date": "2025-01-02" }]`,
- **Then** the original row must be removed,
- **And** 2 new rows must be inserted with the corresponding values.

### 2.2. Empty array

- **Given** a row with `[[payments.id]]`,
- **And** `payments` is `[]`,
- **Then** the row must be removed from the document.

### 2.3. Array key does not exist

- **Given** `[[payments.id]]`,
- **And** `payments` does not exist in `data`,
- **Then** the cell must contain `[[payments.id]]`,
- **And** the row **must not be removed or repeated**.

### 2.4. Array key is not an array

- **Given** `[[user.id]]`,
- **And** `user` exists but is an object (not an array),
- **Then** it must throw an error with a clear message:  
  > "`[[user.id]]` requires 'user' to be an array. Received: [object Object]".

### 2.5. Item property does not exist

- **Given** `[[payments.id]]`,
- **And** an item in `payments` does not have `id`,
- **Then** the cell must contain `[[payments.id]]`.

### 2.6. Item property is `null`/`undefined`

- **Given** `[[payments.amount]]`,
- **And** an item has `amount: null`,
- **Then** the cell must be empty (`""`).

---

## 3. Behavior with formulas and styles

### 3.1. Formulas are preserved and adjusted

- **Given** a cell in a repeatable row with formula `=B3*C3`,
- **When** the row is expanded,
- **Then** each new row must have an adjusted formula: `=B4*C4`, `=B5*C5`, etc.

### 3.2. Styles are copied to new rows

- **Given** a row with cells that have:
  - blue background,
  - thick border,
  - bold font,
- **When** the row is expanded,
- **Then** all those properties must be faithfully copied to the new rows.

### 3.3. Multiple worksheets

- **Given** a workbook with 2 worksheets,
- **And** only the first one contains markers,
- **Then** the second worksheet must remain unchanged.

---

## 4. Coexistence of `{{}}` and `[[]]`

### 4.1. Valid combination

- **Given** a row with `{{user.name}}` and `[[payments.id]]`,
- **And** `payments` is a valid array,
- **Then** the row is repeated N times,
- **And** `{{user.name}}` is resolved against the root data object in each repeated row,
- **And** `[[payments.id]]` is resolved against the current array item.

---

## 5. Input and output

### 5.1. Input as Buffer

- **Given** a valid `.xlsx` file as a `Buffer`,
- **And** a plain `data` object,
- **Then** it must return a `Promise<Buffer>` for the resulting file.

### 5.2. Valid output

- **The resulting Buffer** must open without errors in:
  - Microsoft Excel
  - Google Sheets
  - LibreOffice Calc

---

## 6. Architecture and technical stack

### 6.1. Monorepo with pnpm

- **The package** must live in a workspace managed by `pnpm`.

### 6.2. ESM and CJS

- **It must be distributed** in both formats: ESM and CJS.

### 6.3. ExcelJS dependency

- **ExcelJS** must be a direct dependency (not a peer dependency).
- **It must not require browser-specific polyfills** or depend on a browser environment.

### 6.4. Testing with Vitest

- **All tests** must run with Vitest.
- **Coverage should be ≥90%**.

### 6.5. TypeScript types

- **The API must be fully typed** and generate `.d.ts` files.

---

## 7. Error behavior

### 7.1. Clear error if key is not an array

- **When** `[[]]` is used with a key that is not an array,
- **Then** it must throw an error with context: marker name, received type.

### 7.2. Error on mixed arrays in the same row

- **Given** a row with `[[items.id]]` and `[[payments.id]]`,
- **Then** it must throw an error: "Mixed array keys in row X: items vs payments".

---

## 8. Visual context preservation

### 8.1. Preserve formulas in non-interpolated cells

- **Given** a row with `[[]]` and a cell with formula `=SUM(A:A)`,
- **When** the row is repeated,
- **Then** the formula must be preserved in each new row.

### 8.2. Preserve merges

- **Given** a row with merged cells,
- **When** it is repeated,
- **Then** the new rows should have the same merges.

---

## 9. Public API

### 9.1. Name and signature

- **It must export** `interpolateXlsx(options: { template: Buffer; data: any })`.

### 9.2. Asynchronous

- **It must return** a `Promise<Buffer>`.

---

## 10. Documentation and usage

### 10.1. README must include

- A basic usage example.
- Explanation of `{{}}` vs `[[]]`.
- Behavior with empty arrays and missing keys.
