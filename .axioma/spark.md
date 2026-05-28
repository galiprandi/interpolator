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
