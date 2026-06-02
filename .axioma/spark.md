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

## 2026-05-31 - [Piped Transformations for Template Markers]
**Learning:** Adding a pipe operator (`|`) for string transformations (upper, lower, capitalize, trim, camelCase) provides users with powerful formatting capabilities directly in the template, reducing the need for data preprocessing. Chaining these transforms increases flexibility.
**Pattern:** Decouple path resolution from value transformation. Resolve the value first, then apply a series of transformations. This allows the same transformation logic to be reused across different marker types (interpolation and array expansion) and data sources (main data and fallbacks).

## 2026-06-20 - [Non-string Transformations and JSON Formatting]
**Learning:** Limiting transformations to strings is an unnecessary constraint that limits the utility of the pipe operator. By allowing any value to enter the transformation chain and providing a `json` transform, users can debug their data or output complex objects as strings directly in the template.
**Pattern:** Remove early type-exits in transformation utilities. Guard string-specific logic with type checks and add a universal `json` (stringification) transform to handle non-primitive data gracefully.

## 2026-06-25 - [Utility-Driven Data Presentation]
**Learning:** Users often need minor data formatting (joining lists, pluralization, rounding numbers) that doesn't justify a new dependency or complex backend preprocessing. Adding lightweight, type-guarded transforms directly to the core resolution engine provides high value with negligible overhead.
**Pattern:** Implement "micro-transforms" like `join`, `plural`, and `round` with defensive type checking. This ensures the template engine remains robust while offering flexible presentation logic for various data types (arrays, numbers, strings).
