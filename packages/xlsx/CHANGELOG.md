# Changelog

All notable changes to `@interpolator/xlsx` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Playground support via `apps/playground` to experiment with XLSX templates (invoice and invoice-visual examples).
- Additional functional tests covering edge-cases for missing markers, null values, styles and multi-worksheet scenarios.

### Fixed
- Improved handling of row expansion with merged cells by ensuring merges are replicated from the template row after insertion.
- Clarified and aligned behavior for array markers `[[array.key]]` and root markers `{{path.to.key}}` across tests and documentation.

## [0.1.0] - 2024-01-01

### Added
- Initial release of `@interpolator/xlsx` with support for:
  - Single value interpolation using `{{path.to.key}}`.
  - Row repetition for arrays using `[[array.key]]`.
  - Basic preservation of formatting, formulas and merges from the template workbook.
