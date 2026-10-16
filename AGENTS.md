# AI Agent & Contributor Guide — php-excel-helper

This guide defines core conventions, architecture, and verification commands for contributors and AI coding agents working on `phatnt93/php-excel-helper`.

## Project Overview

- **Package Name:** `phatnt93/php-excel-helper` (namespace `ExcelHelper\`)
- **Target Runtime:** PHP `^8.1` (tested across 8.1, 8.2, 8.3, 8.4)
- **Runtime Dependency:** `phpoffice/phpspreadsheet` `^5.8` (5.9+ requires PHP 8.2; only use APIs available in 5.8.0)
- **Primary Goal:** Read and write spreadsheets (xlsx, xls, ods, csv) through plain array configuration — start row/column, typed columns with defaults and validation, row-level error reporting, styled exports — with a parallel fluent builder API.
- **Reference Standards:** PER Coding Style 2.0 / PSR-12, PSR-4 autoloading (`src/` and `tests/`).

## Architecture Map (actual)

```
src/
├── Excel.php        # Static facade: read, reader, chunk, write, writer, registerType, setLocale
├── Contracts/       # TypeInterface (two-way cell type: read + write)
├── Config/          # Array config parsers → immutable value objects, ConfigPath for error messages
├── Reader/          # ReadOptions, ExcelReader, ColumnResolver, RowMapper, ReadResult, RowError, chunking
├── Writer/          # WorkbookDefinition, ExcelWriter, SheetRenderer, CellWriter, StyleMapper, ExcelOutput
├── Types/           # Built-in types + TypeRegistry
├── Support/         # Coordinate, Color, SizeParser, DownloadHeaders
├── I18n/            # Translator and dictionaries (en default, vi)
└── Exceptions/      # ExcelHelperException interface + concrete exceptions
```

## Core Rules & Conventions

1. **Strict Typing:** Every PHP file in `src/`, `tests/`, and `examples/` MUST start with `declare(strict_types=1);`.
2. **Coding Style:** PER-CS 2.0 enforced by PHP-CS-Fixer (`.php-cs-fixer.dist.php`).
3. **Static Analysis:** PHPStan level 8 (`phpstan.neon.dist`, `phpVersion: 80100`).
4. **Test Coverage:** line coverage ≥ 90% (`tests/coverage-check.php`). Every type needs a write → read round-trip test.
5. **Config vs data errors:** invalid configuration throws `InvalidConfigException` with a config path (e.g. `columns.C.type`); invalid cell data produces `RowError` entries according to `on_error`.
6. **Array/builder parity:** array config and fluent builders must produce the same value objects and results.
7. **Security-First Rules:**
   - Detect the real format by content (`IOFactory::identify`) and enforce the `xlsx/xls/ods/csv` whitelist.
   - Do not run formulas from read files by default; use cached values.
   - Strings are always written as `DataType::TYPE_STRING`; formulas only with `type = formula`; CSV escapes leading `= + - @ \t \r`.
   - All regexes with end anchor `$` must use the `D` flag or `\z`.
   - Wrap PhpSpreadsheet exceptions in library exceptions (keep `previous`).
8. **Internationalization (i18n):** default locale `en` with fallback to `en`; developer exception messages in English; end-user data errors via `Translator` keys — update both `src/I18n/en.php` and `src/I18n/vi.php`.
9. **Examples:** every script in `examples/` is self-contained, creates its own data, runs with `php examples/xx.php`, exits with code 0.

## Verification Commands

```bash
composer validate --strict
composer cs:check
composer analyse
composer test
composer test:coverage
for f in examples/*.php; do php "$f"; done
```
