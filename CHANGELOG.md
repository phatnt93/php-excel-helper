# Changelog

All notable changes to `phatnt93/php-excel-helper` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-10-16

### Added
- **Template Engine (`Excel::template()`):**
  - Loads existing Excel template workbooks, preserving existing styles, fonts, formulas, and layouts.
  - Replaces placeholders (`{{key}}` or `{{ key }}`) across the spreadsheet with direct value and type mapping (`set()`, `setValues()`).
  - Inserts dynamic table data with automatic downward shifting of subsequent rows and style inheritance (`insertTable()`).
- **Image & Logo Insertion (`ImageDefinition`):**
  - Inserts logos and images into worksheets with custom coordinates, width, height, aspect ratio control, and pixel offsets.
  - Configurable in array configuration (`images => [...]`) or via fluent builders (`image()`, `images()`).
- **Data Validation & Dropdown Lists:**
  - Adds cell/column data validation rules (`validation => [...]`): `list` (dropdown selection), `whole` (integer range), `decimal`, and `text_length`.
  - Configurable error alerts, custom titles, and messages directly in Excel.
- **Conditional Formatting:**
  - Sheet-level (`conditional_formatting => [...]`) and column-level (`conditional => [...]`) rules.
  - Supports operators (`<`, `<=`, `>`, `>=`, `==`, `!=`, `between`), custom font colors, bold/italic, and background fill styling.
- **Fast Streaming Reader (`FastXlsxStreamReader`, `CsvStreamReader::stream`, `Excel::stream()`):**
  - Pull-based streaming XMLReader over ZipArchive for low-memory row streaming of large XLSX spreadsheets and lazy O(1) memory streaming for UTF-8 CSV files.
  - Provides `Excel::stream()` generator for reading XLSX and CSV row-by-row with minimal RAM usage.
- New standalone demonstration example `examples/07_advanced_v1_1_features.php`.
- Test suite expanded to 271 tests, 924 assertions, 100% passing, PHPStan level 8 clean, CS-Fixer clean, line coverage >= 90%.
## [1.0.0] - 2026-10-16

### Added
- Core `Excel` static facade with `read()`, `reader()`, `write()`, `writer()`, `chunk()`, `sheetNames()`, `info()`, `headers()`, `setLocale()`, and `registerType()`.
- Full support for `xlsx`, `xls`, `ods`, and `csv` formats with automatic content-based format inspection and strict format whitelisting.
- Two-way extensible type system (`TypeInterface`, `TypeRegistry`, `CellValue`, `ColumnContext`):
  - Built-in types: `string`, `int`, `float`, `number`, `decimal`, `bool`, `date`, `datetime`, `time`, `raw`, `formatted`, `auto`, `percent`, `money`, `formula`, `link`.
  - Seamless handling of Excel date serial numbers across 1900 and 1904 calendar systems.
  - Strict preservation of phone numbers and string values with leading zeros.
- 3 intuitive column declaration forms for reading:
  - (a) Keyed by column letter or 1-based number (`'A' => 'code'`, `'C' => ['key' => 'name']`).
  - (b) Sequential column list starting from `start_column` with `null` column skipping.
  - (c) Header text matching with case-insensitivity, whitespace trimming, and Unicode NFC normalization.
- Cell and row processing pipeline: `trim` -> `required` / `default` -> `map` / `map_strict` -> `Type::read` -> `transform` -> `validate`.
- Multilingual i18n row error reporting in English (`en`) and Vietnamese (`vi`) with 3 error handling policies: `collect`, `skip_row`, `throw`.
- Large file chunking engine (`Excel::chunk`) and constant O(1) memory CSV streaming reader (`CsvStreamReader`).
- Styled workbook exports:
  - Title banner merging across columns.
  - Flat, user-friendly style mapper (font, fill/background, alignment, borders).
  - Bulk number formatting and automatic totals calculation formulas (`sum`, `average`, `min`, `max`, `count`).
  - Zebra striping, freeze panes, auto filter, row height, and custom merges.
  - Direct worksheet access via `tap` hook.
- Output management (`ExcelOutput`): atomic file save, direct download with RFC 5987 UTF-8 headers, binary string generation, and raw PhpSpreadsheet access.
- Security-first hardening:
  - Formula injection mitigation via explicit `TYPE_STRING` in Excel and escaping in CSV.
  - File format spoofing protection and XXE prevention.
  - Safety limits for maximum file size and maximum rows.
  - Hyperlink scheme whitelisting (`http`, `https`, `mailto`).
- Fluent builders (`ReaderBuilder`, `WriterBuilder`, `SheetBuilder`) with 100% configuration parity with array options.
- 6 standalone example scripts in `examples/`.
- Bilingual documentation (`README.md`, `README.vi.md`) and AI agent context files (`llms.txt`, `llms-full.txt`).
- Complete automated test suite: 236 tests, 779 assertions, 100% passing, 0 risky/warnings, PHPStan level 8 clean, line coverage >= 90%.
