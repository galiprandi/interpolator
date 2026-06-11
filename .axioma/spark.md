# Spark's Journal

## 2025-05-15 - [XLSX Template Enhancements]
**Learning:** Users often need to render simple lists or include row counters in their documents. Adding support for primitive arrays and special index markers significantly improves the library's versatility with minimal code changes.
**Pattern:** Using special property paths (like `$index`, `$number`) is an effective way to expose metadata about the current iteration without cluttering the input data object.

## 2025-05-16 - [Smart Type Preservation & Dynamic Metadata]
**Learning:** Automatically preserving data types (like Dates and Numbers) when a cell contains only a marker is much more intuitive than forcing everything to strings. Also, interpolating worksheet names is a frequently requested feature for multi-sheet reports.
**Pattern:** Detect "single-marker cells" using regex anchor patterns (`/^...$/`) to decide between direct value assignment (preserving type) and string replacement.

## 2025-05-17 - [Enhanced Array Metadata]
**Learning:** Providing boolean flags like `$first` and `$last` allows users to implement conditional formatting or separators without complex logic in the data source.
**Pattern:** Extend existing metadata marker logic to support Booleans and array-level metrics (like `$length`) using the same path-based resolution pattern.

## 2025-05-18 - [Styling and Conditional Markers]
**Learning:** Adding `$even` and `$odd` markers makes it trivial for users to implement zebra-striping or alternating layouts in Excel without any preprocessing on the data side.
**Pattern:** Always think of how metadata markers can replace manual data enrichment. If a value can be derived from the iteration context (like parity), it's a candidate for a metadata marker.

## 2025-05-19 - [Excel Coordinate Markers]
**Learning:** Providing markers that translate numeric indices into Excel coordinates (like `$colLetter` or `$cell`) bridges the gap between structured data and document layout. This allows users to create more dynamic references or labels within the template without needing to pre-calculate Excel-specific addresses in their data source.
**Pattern:** Implement a robust utility for coordinate translation (e.g., column index to letter) and expose it through both global and iteration contexts to ensure consistency.

## 2024-05-22 - Excel Context Markers
**Learning:** Contextual information (, , , ) is frequently requested in document templates but often requires the user to manually enrich their data object. Providing these as built-in markers dramatically simplifies report generation templates.
**Pattern:** Abstract path resolution into a context-aware helper that wraps the standard data resolution. This allows for clean injection of environment/runtime variables without polluting the user's data object.

## 2026-05-29 - [Sheet Metadata and Root Context Consistency]
**Learning:** Exposing workbook-level metadata (like sheet count and position) allows for more professional, context-aware reports (e.g., "Sheet 1 of 5"). Aligning root context markers with iteration markers (like parity) creates a more predictable API for users.
**Pattern:** Ensure all available document hierarchy metadata is injectable into the resolution context. Consistency between global markers and loop-local markers reduces cognitive load for template authors.

## 2026-06-05 - [Comprehensive Metadata Markers & Descriptive Aliases]
**Learning:** Users have different preferences for indexing (0-based vs 1-based) and naming conventions (e.g., $col vs $columnLetter). Providing a broad set of descriptive aliases and date components ($year, $month, $day) makes the template syntax more expressive and reduces the need for external data manipulation.
**Pattern:** Identify and implement common aliases and derived metadata (like date components or column-based parity) to cater to diverse developer preferences and use cases, ensuring the template feels "native" to both Excel and standard programming models.

## 2026-06-10 - [Roadmap: @interpolator/markdown]
**Learning:** There is strong interest in extending the interpolation concept to Markdown. This requires handling different structural constraints (like table syntax and indentation) compared to XLSX.
**Pattern:** Future packages should follow the same declarative syntax ({{}} and [[]]) to ensure a consistent experience across all `@interpolator` formats.

## 2026-06-15 - [Centralized Marker Resolution and Advanced Context]
**Learning:** Consolidating marker resolution into a single context-aware function ('resolveInternal') significantly simplifies the interpolation engine. It allows for consistent behavior between root-level and array-level markers, and makes it trivial to add new metadata markers like $hour or $isHeader.
**Pattern:** Pass a rich 'context' object (containing sheet, row, col, index, length, etc.) through the resolution chain to enable powerful, context-sensitive markers without complicating the user's data schema.

## 2026-07-05 - [Advanced Collection and String Transforms]
**Learning:** For template authors, basic collection operations like `reverse`, `sort`, and `compact` are essential for data presentation without requiring backend changes. Additionally, providing `sum` and `avg` aggregations directly in the template allows for quick report generation from raw data.
**Pattern:** When implementing string/array transformations, always use non-mutating methods (like `[...arr].reverse()`) and ensure Unicode safety for strings using the spread operator. Robust numeric aggregations should gracefully handle non-numeric data to prevent template crashes.

## 2026-05-31 - [Piped Transformations for Template Markers]
**Learning:** Adding a pipe operator (`|`) for string transformations (upper, lower, capitalize, trim, camelCase) provides users with powerful formatting capabilities directly in the template, reducing the need for data preprocessing. Chaining these transforms increases flexibility.
**Pattern:** Decouple path resolution from value transformation. Resolve the value first, then apply a series of transformations. This allows the same transformation logic to be reused across different marker types (interpolation and array expansion) and data sources (main data and fallbacks).

## 2026-06-20 - [Non-string Transformations and JSON Formatting]
**Learning:** Limiting transformations to strings is an unnecessary constraint that limits the utility of the pipe operator. By allowing any value to enter the transformation chain and providing a `json` transform, users can debug their data or output complex objects as strings directly in the template.
**Pattern:** Remove early type-exits in transformation utilities. Guard string-specific logic with type checks and add a universal `json` (stringification) transform to handle non-primitive data gracefully.

## 2026-06-25 - [Utility-Driven Data Presentation]
**Learning:** Users often need minor data formatting (joining lists, pluralization, rounding numbers) that doesn't justify a new dependency or complex backend preprocessing. Adding lightweight, type-guarded transforms directly to the core resolution engine provides high value with negligible overhead.
**Pattern:** Implement "micro-transforms" like `join`, `plural`, and `round` with defensive type checking. This ensures the template engine remains robust while offering flexible presentation logic for various data types (arrays, numbers, strings).

## 2026-06-30 - [Numeric Math Transforms]
**Learning:** Providing basic math operations (floor, ceil, abs) alongside rounding allows users to handle financial or statistical data presentation directly in templates without upstream modification.
**Pattern:** Use type-guarded `Math` function wrappers to extend the transformation engine safely for numeric data types.

## 2025-05-20 - [Nested Array Resolution & Multi-line Row Expansion]
**Learning:** Decoupling array identification from simple root keys enables much more powerful templates. Supporting nested paths (e.g., `[[order.items.name]]`) and transformations (e.g., `[[items | reverse]]`) in expansion markers allows template authors to reshape data for display without backend changes. Adding a `lines` transform to split strings into arrays further bridges the gap between raw text data and tabular document structure.
**Pattern:** Implement a path-probing utility (like `findArrayInPath`) that iteratively checks segments of a path to find the collection context. This provides a natural, intuitive syntax for nested data while maintaining backward compatibility.

## 2026-07-10 - [Numeric Extremes and Nullability]
**Learning:** When implementing numeric aggregations like `min` and `max`, returning `undefined` for empty or non-numeric collections is more accurate than defaulting to `0`. This allows template authors to use default values (e.g., `{{ items | min || N/A }}`) to handle missing data explicitly.
**Pattern:** Prefer `undefined` over fallback values for collection reductions when the identity element (like 0 for sum) could be misinterpreted as a valid result from the data.

## 2026-07-20 - [Unicode-Safe String Transforms]
**Learning:** When extracting characters from strings (like for an `initials` transform), standard indexing (`str[0]`) can break on Unicode surrogate pairs like emojis. Using the spread operator (`[...str][0]`) ensures that the transform is robust and "future-proof" for internationalized data.
**Pattern:** Always use Unicode-aware character access in string utilities to maintain the "high-quality, dependency-free" standard of the library.

## 2026-07-15 - [Object and Array Manipulation Transforms]
**Learning:** Providing basic object introspection (`keys`, `values`) and array manipulation (`flat`) transforms allows users to handle more complex data structures directly in templates. Updating `length` to support objects makes the API more consistent across different data types.
**Pattern:** Extend core utilities to support both arrays and objects where it makes sense (like `length`), and provide specific transforms for type-specific operations that follow standard JavaScript behavior.

## 2026-07-25 - [Sequence Generation via Range Transform]
**Learning:** Providing a way to generate sequences directly from a number (via `range`) allows users to create dynamic row counts or lists without needing to pre-calculate arrays in their data source. This is particularly useful for things like "Top N" lists or fixed-size forms.
**Pattern:** Use simple, numeric-to-collection transforms to bridge the gap between scalar data and structural document requirements (like row expansion).

## 2025-05-21 - [Object Entry Transformation for Iteration]
**Learning:** Enabling iteration over object properties is a common requirement for dynamic templates. By providing an `entries` transform that converts objects into a standardized `{ key, value }` array, we bridge the gap between static objects and the array-based row expansion engine.
**Pattern:** Map native language structures (like `Object.entries`) into a format that fits the library's existing iteration patterns (array of objects), ensuring consistency between different data shapes and their visual representation in the document.
