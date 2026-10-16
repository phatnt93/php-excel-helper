# PHP Excel Helper

[![Latest Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/phatnt93/php-excel-helper)
[![PHP Version](https://img.shields.io/badge/php-%5E8.1-777bb4.svg)](https://php.net)
[![PHPStan](https://img.shields.io/badge/PHPStan-level%208-brightgreen.svg)](https://phpstan.org)
[![Tests](https://img.shields.io/badge/tests-100%25%20passing-brightgreen.svg)](https://phpunit.de)
[![Coverage](https://img.shields.io/badge/coverage-%E2%89%A590%25-brightgreen.svg)](https://github.com/phatnt93/php-excel-helper)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Read and write Excel and CSV spreadsheets (`.xlsx`, `.xls`, `.ods`, `.csv`) in PHP with simple, intuitive array configuration. Built on top of [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet).

**[Đọc tài liệu Tiếng Việt (Vietnamese README)](README.vi.md)**

---

## Table of Contents

- [Why php-excel-helper?](#why-php-excel-helper)
- [Installation](#installation)
- [Quick Start](#quick-start)
  - [Read to Array](#read-to-array)
  - [Write from Array](#write-from-array)
  - [Fluent Builders](#fluent-builders)
- [Reading Spreadsheets](#reading-spreadsheets)
  - [Options Reference](#read-options-reference)
  - [3 Ways to Define Columns](#3-ways-to-define-columns)
  - [Column Attributes](#read-column-attributes)
  - [Supported Data Types](#supported-data-types)
  - [Error Handling (`on_error`)](#error-handling-on_error)
- [Writing Spreadsheets](#writing-spreadsheets)
  - [Workbook Options](#workbook-options)
  - [Sheet Options](#sheet-options)
  - [Friendly Style System](#friendly-style-system)
  - [Output Targets (Save, Download, String)](#output-targets)
- [Large Files & Memory Management](#large-files--memory-management)
  - [Chunk Reading](#chunk-reading)
  - [Zero-Memory CSV Streaming](#zero-memory-csv-streaming)
- [Utilities](#utilities)
- [Security First](#security-first)
- [Custom Data Types](#custom-data-types)
- [How-To Recipes (FAQ)](#how-to-recipes-faq)
- [License](#license)

---

## Why php-excel-helper?

[PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) is the industry standard for spreadsheet processing in PHP, but its low-level API requires dozens of lines of boilerplate for everyday tasks: manually iterating cells, converting Excel serial dates, handling phone numbers with leading zeros, building multi-level style arrays, and configuring chunk filters.

`php-excel-helper` provides an **elegant, opinionated layer** above PhpSpreadsheet:

| Feature | Raw PhpSpreadsheet | php-excel-helper |
|---|---|---|
| **API Style** | Low-level OO, cell-by-cell loops | Array-first configuration or fluent builders |
| **Excel Dates** | Float serial (e.g. `46311.5`) requires manual calendar conversion | Auto-converted to `Y-m-d` or `DateTimeImmutable` |
| **Phone Numbers / Leading Zeros** | Cast to integer `901234567` by default | Preserved as string `'0901234567'` |
| **Row Validation & Errors** | Manual cell validation | Built-in row errors (`required`, `invalid_int`, `custom`, etc.) |
| **Formula Injection** | Must manually sanitize text | Auto-escaped (`TYPE_STRING` in Excel, escaped in CSV) |
| **Styling** | 5-level nested arrays (`['font' => ['bold' => true]]`) | Flat friendly keys (`'bold' => true, 'bg' => '#1F4E78'`) |
| **Large CSVs** | Loads entire file into memory (high RAM) | Streamed with `SplFileObject` in O(1) RAM |
| **Direct Escape** | Hard to integrate outside standard API | Escape via `->spreadsheet()` or `tap($sheet)` |

---

## Installation

```bash
composer require phatnt93/php-excel-helper
```

**Requirements:**
- PHP 8.1 or higher
- Extensions: `mbstring` (standard)
- Suggested: `ext-intl` (for Unicode normalization when matching column headers)

---

## Quick Start

### Read to Array

```php
use ExcelHelper\Excel;

$result = Excel::read('users.xlsx', [
    'header_row' => 1,
    'columns' => [
        ['header' => 'Employee ID', 'key' => 'code'],
        ['header' => 'Full Name',   'key' => 'name'],
        ['header' => 'Date of Birth','key' => 'birthday', 'type' => 'date'],
        ['header' => 'Salary',      'key' => 'salary',   'type' => 'number'],
        ['header' => 'Active',      'key' => 'active',   'type' => 'bool'],
    ],
]);

// Returns array of clean associative arrays
foreach ($result->rows() as $row) {
    echo "{$row['code']}: {$row['name']} ({$row['salary']})\n";
}

// Any data errors are captured cleanly
if ($result->hasErrors()) {
    foreach ($result->errors() as $error) {
        echo "Row {$error->row}, column {$error->column} ({$error->key}): {$error->message}\n";
    }
}
```

### Write from Array

```php
use ExcelHelper\Excel;

Excel::write([
    'title' => 'MONTHLY PAYROLL REPORT',
    'columns' => [
        'code'   => 'Code',
        'name'   => ['header' => 'Employee Name', 'width' => 28],
        'salary' => ['header' => 'Salary', 'type' => 'number', 'format' => '#,##0'],
        'active' => ['header' => 'Status', 'type' => 'bool', 'labels' => [true => 'Active', false => 'Inactive']],
    ],
    'data' => [
        ['code' => 'NV001', 'name' => 'Alice Johnson', 'salary' => 4500, 'active' => true],
        ['code' => 'NV002', 'name' => 'Bob Smith',     'salary' => 5200, 'active' => false],
    ],
    'totals' => ['salary' => 'sum'],
    'freeze_header' => true,
    'zebra' => true,
])->save('payroll.xlsx');
```

### Fluent Builders

Every array option has a matching fluent builder method:

```php
// Read using builder
$result = Excel::reader('users.xlsx')
    ->headerRow(1)
    ->column('A', 'code')
    ->column('B', 'name')
    ->column('C', 'salary', type: 'number')
    ->onError('skip_row')
    ->read();

// Write using builder
Excel::writer()
    ->title('Employees')
    ->column('code', 'Code')
    ->column('name', 'Name', width: 25)
    ->data($employees)
    ->save('output.xlsx');
```

---

## Reading Spreadsheets

### Read Options Reference

| Option | Type | Default | Description |
|---|---|---|---|
| `format` | `string` | `'auto'` | Format to use (`'auto'`, `'xlsx'`, `'xls'`, `'ods'`, `'csv'`). Auto sniffs by file signature. |
| `allowed_formats` | `list<string>` | `['xlsx','xls','ods','csv']` | Whitelist of allowed formats. |
| `sheet` | `int\|string` | `0` | Sheet 0-based index or string tab name. |
| `header_row` | `?int` | `null` | 1-based header row number. When set, enables matching columns by header text. |
| `start_row` | `int` | `header_row + 1` or `1` | 1-based first row containing data. |
| `end_row` | `?int` | `null` | 1-based last data row (`null` = read to last row). |
| `columns` | `array` | `[]` | Column definitions. Empty = read all columns as raw strings. |
| `skip_empty_rows` | `bool` | `true` | Skip rows where all mapped cells are blank. |
| `stop_at_empty_row`| `bool` | `false` | Stop reading when encountering the first completely blank row. |
| `trim` | `bool` | `true` | Trim whitespace from string cell values. |
| `on_error` | `string` | `'collect'` | Error handling mode: `'collect'`, `'skip_row'`, or `'throw'`. |
| `row_number_key` | `?string` | `null` | If set (e.g. `'_row'`), includes the 1-based Excel row number in each returned row. |
| `calculate_formulas`| `bool` | `false` | `false` = read cached calculated value; `true` = evaluate formula. |
| `max_rows` | `?int` | `null` | Safety limit: aborts if file exceeds row limit (checked before load). |
| `max_file_size` | `int\|string\|null`| `null` | File size limit (`'20M'`, `'512K'`, or bytes). |
| `date_format` | `?string` | `'Y-m-d'` | Default date string format (`null` = returns `DateTimeImmutable`). |
| `datetime_format` | `?string` | `'Y-m-d H:i:s'` | Default datetime format. |
| `csv` | `array` | `[]` | CSV options: `delimiter` (default `auto`), `enclosure`, `input_encoding` (`auto`). |
| `locale` | `?string` | `null` | Error message translation locale (`'en'`, `'vi'`). |

---

### 3 Ways to Define Columns

#### (a) Keyed by Column Letter or Number
```php
'columns' => [
    'A' => 'code',             // Shorthand for ['key' => 'code', 'type' => 'string']
    'B' => ['key' => 'name'],
    3   => ['key' => 'salary', 'type' => 'float'], // 1-based index (3 = C)
]
```

#### (b) Sequential List from `start_column`
```php
'start_column' => 'B',
'columns' => [
    'code',                     // B
    'name',                     // C
    null,                       // D is skipped
    ['key' => 'salary', 'type' => 'float'], // E
]
```

#### (c) Matched by Header Text
```php
'header_row' => 1,
'columns' => [
    ['header' => 'Employee ID', 'key' => 'code'],
    ['header' => ['Email', 'E-mail Address'], 'key' => 'email', 'optional' => true],
]
```

---

### Read Column Attributes

| Attribute | Type | Default | Description |
|---|---|---|---|
| `key` | `string` | — | Required. Output array key. |
| `type` | `string` | `'string'` | Data type (`'string'`, `'int'`, `'float'`, `'decimal'`, `'bool'`, `'date'`, `'datetime'`, `'time'`, `'raw'`, `'formatted'`). |
| `format` | `?string` | by type | Output date format (e.g. `'d/m/Y'`) or decimal places. |
| `input_format` | `string\|list<string>` | `null` | Format(s) for parsing date strings (e.g. `['d/m/Y', 'Y-m-d']`). |
| `default` | `mixed` | `null` | Fallback value when cell is empty or on error. |
| `required` | `bool` | `false` | When true, empty cell records a `required` error. |
| `trim` | `?bool` | by global | Override trimming for this column. |
| `map` | `?array` | `null` | Dictionary translation before casting (e.g. `['Nam' => 'male', 'Nữ' => 'female']`). |
| `map_strict` | `bool` | `false` | Record a `not_in_map` error if value is not in dictionary. |
| `transform` | `?callable` | `null` | Custom transformation callback: `fn(mixed $val, RowContext $ctx): mixed`. |
| `validate` | `?callable` | `null` | Custom validation: `fn(mixed $val, RowContext $ctx): true|string`. |
| `optional` | `bool` | `false` | For header matching: column may be absent in file without error. |

---

### Supported Data Types

- **`string`**: Preserves leading zeros (`'0901234567'`). Converts scientific notation floats (`1.0E+15`) to exact integers.
- **`int`**: Numeric string or float without fractional part (`42.0` -> `42`). Supports `thousands_separator`.
- **`float`** (alias **`number`**): Standard float. Configurable `thousands_separator` and `decimal_separator`.
- **`decimal`**: Returns a string representation preserving exact decimal precision (for currency/finance).
- **`bool`**: `true/false`, `1/0`, and case-insensitive strings (`'yes'/'no'`, `'true'/'false'`, `'đúng'/'sai'`).
- **`date`**: Parses Excel serial numbers (1900 and 1904 calendars) or date strings. Returns formatted string or `DateTimeImmutable`. Rejects overflow dates like `31/02/2026`.
- **`datetime`**: Full date and time.
- **`time`**: Time fraction or time string (e.g. `'14:30:00'`).
- **`raw`**: Value as returned by PhpSpreadsheet without casting.
- **`formatted`**: Cell display string formatted exactly as visible in Excel.

---

### Error Handling (`on_error`)

- **`collect` (default)**: Records errors in `$result->errors()`, fills default value, and **keeps** the row.
- **`skip_row`**: Records errors and **omits** the invalid row from `$result->rows()`.
- **`throw`**: Immediately throws `RowValidationException` on the first error.

```php
$result = Excel::read('users.xlsx', [
    'on_error' => 'collect',
    'columns' => [
        'A' => ['key' => 'age', 'type' => 'int', 'required' => true, 'default' => 0],
    ],
]);

if ($result->hasErrors()) {
    $errorsByRow = $result->errorsByRow(); // array<int, list<RowError>>
    foreach ($errorsByRow as $rowNumber => $rowErrors) {
        foreach ($rowErrors as $err) {
            echo "Line {$rowNumber}: {$err->message}\n";
        }
    }
}
```

---

## Writing Spreadsheets

### Workbook Options

```php
Excel::write([
    'format'     => 'xlsx', // 'xlsx', 'xls', 'ods', 'csv'
    'properties' => [
        'creator'     => 'Acme Inc',
        'title'       => 'Annual Report',
        'description' => 'Summary of quarterly metrics',
    ],
    'defaults'   => [
        'zebra'         => true,
        'freeze_header' => true,
    ],
    'sheets'     => [
        ['name' => 'Q1', 'columns' => [...], 'data' => [...]],
        ['name' => 'Q2', 'columns' => [...], 'data' => [...]],
    ],
])->save('report.xlsx');
```

---

### Sheet Options

- `name`: Tab title (sanitized against Excel's 31-character limit and forbidden characters).
- `title`: Large banner title above the table, merged across columns.
- `start_cell`: Top-left origin cell (default `'A1'`).
- `columns`: Column definitions (keys, headers, types, widths, styles).
- `data`: Iterable rows (`array`, `Traversable`, or `Generator`).
- `totals`: Totals row formulas (`['revenue' => 'sum', 'rating' => 'average']`).
- `zebra`: Alternating row styling (`true` or `['color' => '#F2F2F2']`).
- `borders`: Table borders (`true`, `'double'`, `'thin'`, or custom array).
- `freeze_header`: Freeze header row so it stays visible while scrolling.
- `auto_filter`: Add Excel dropdown filters to header row.
- `auto_size`: Auto calculate column width (automatically disabled if sheet exceeds `5000` rows).
- `tap`: Callback to interact with raw `Worksheet` and `SheetLayout`.

---

### Friendly Style System

```php
Excel::write([
    'header_style' => [
        'bold'       => true,
        'color'      => '#FFFFFF',
        'background' => '#1F4E78',
        'align'      => 'center',
        'valign'     => 'middle',
    ],
    'columns' => [
        'code'   => ['header' => 'Code', 'align' => 'center'],
        'amount' => ['header' => 'Amount', 'style' => ['bold' => true]],
    ],
    'borders' => 'thin',
    'zebra'   => true,
    'data'    => $rows,
])->save('styled.xlsx');
```

---

### Output Targets

```php
$output = Excel::write($config);

// 1. Save to local path
$output->save('/exports/file.xlsx');

// 2. Stream download with RFC 5987 UTF-8 headers
$output->download('báo-cáo.xlsx');

// 3. Get raw binary string (for S3, email attachments, queues)
$binary = $output->toString();

// 4. Escape hatch to native PhpSpreadsheet instance
$spreadsheet = $output->spreadsheet();
```

---

## Large Files & Memory Management

### Chunk Reading

Process huge spreadsheets without loading the entire document into RAM:

```php
$summary = Excel::chunk('huge.xlsx', ['columns' => [...]], 1000,
    function (array $chunkRows, ChunkContext $ctx): void {
        echo "Processing chunk #{$ctx->chunkIndex()} ({$ctx->rowCount} rows)...\n";

        // Insert batch into database
        DB::table('users')->insert($chunkRows);

        // Optional early stopping:
        // if ($someCondition) { return false; }
    }
);

echo "Total rows processed: {$summary->totalRows()}\n";
```

### Zero-Memory CSV Streaming

When reading CSV files in chunk mode, `php-excel-helper` bypasses PhpSpreadsheet entirely and streams lines via `SplFileObject`. This guarantees **constant O(1) memory usage** even on files with 500,000+ lines.

---

## Utilities

```php
// List sheet names without loading cell data
$sheets = Excel::sheetNames('workbook.xlsx');

// Inspect sheet dimensions
$info = Excel::info('workbook.xlsx');
foreach ($info as $sheetInfo) {
    echo "{$sheetInfo->name}: {$sheetInfo->totalRows} rows, {$sheetInfo->totalColumns} columns\n";
}

// Read header row directly
$headers = Excel::headers('workbook.xlsx', row: 1);
```


---

## Template Engine (`Excel::template`)

Fill pre-formatted Excel templates (invoices, contracts, certificates) while preserving branding, logos, cell formats, and complex formulas:

```php
Excel::template('templates/invoice_template.xlsx')
    ->set('customer_name', 'Acme Corp')
    ->set('invoice_no', 'INV-2026-001')
    ->setValues([
        'date' => '2026-10-16',
        'due_date' => '2026-11-16',
        'B15' => 'Direct cell address filling',
    ])
    ->insertTable('item_name', [
        ['item_name' => 'Web Hosting Pro', 'qty' => 1, 'price' => 2500000],
        ['item_name' => 'Domain .com', 'qty' => 2, 'price' => 750000],
    ])
    ->save('exports/invoice_001.xlsx');
```

* Sub-rows in `insertTable()` duplicate formatting and automatically shift footer rows and total formulas downward!

---

## Advanced v1.1.0 Features

### 1. Fast Streaming Generator (`Excel::stream`)

Read ultra-large XLSX and CSV spreadsheets row-by-row as a PHP `Generator` with constant O(1) memory:

```php
foreach (Excel::stream('huge_file.xlsx', ['start_row' => 2]) as $rowNumber => $row) {
    // Process single row with negligible memory overhead
}
```

### 2. Image & Logo Insertion

Insert corporate logos and product thumbnails with coordinates and pixel dimensions:

```php
Excel::write([
    'images' => [
        ['path' => 'assets/logo.png', 'cell' => 'A1', 'width' => 120, 'height' => 50],
    ],
    'columns' => [...],
    'data' => [...],
])->save('report.xlsx');
```

### 3. Data Validation & Dropdowns

Add Excel dropdown selection lists and validation rules directly to columns:

```php
'status' => [
    'header' => 'Status',
    'validation' => ['Approved', 'Pending', 'Rejected'], // Dropdown list
],
'age' => [
    'header' => 'Age',
    'validation' => ['type' => 'integer', 'min' => 18, 'max' => 65],
],
```

### 4. Conditional Formatting

Highlight values automatically with Excel conditional styles:

```php
'balance' => [
    'header' => 'Balance',
    'conditional' => [
        ['operator' => '<', 'value' => 0, 'color' => 'FF0000', 'bg' => 'FFE0E0', 'bold' => true],
    ],
],
```
---

## Security First

1. **Formula Injection Guard**: Untrusted text fields starting with `=`, `+`, `-`, `@`, `\t`, `\r` are explicitly stored as text (`TYPE_STRING`) in Excel exports, and quoted in CSV exports.
2. **Format Spoofing Protection**: Files are identified by binary magic bytes and content, strictly matching against the allowed whitelist. HTML or script files disguised as `.xlsx` are rejected.
3. **XXE Protection**: Relies on PhpSpreadsheet's XML scanner to neutralize XML External Entity attacks.
4. **Link Sanitization**: Hyperlinks only accept safe schemes (`http`, `https`, `mailto`); `javascript:` links are rejected.
5. **Memory Cap**: `max_file_size` and `max_rows` prevent memory exhaustion before loading.

---

## Custom Data Types

Register custom bidirectional data types:

```php
use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Types\CellValue;
use ExcelHelper\Types\ColumnContext;

class MoneyVndType implements TypeInterface
{
    public function name(): string
    {
        return 'money_vnd';
    }

    public function read(mixed $value, ColumnContext $context): float
    {
        return (float) preg_replace('/[^\d]/u', '', (string) $value);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        return CellValue::numeric((float) $value, '#,##0 "đ"');
    }
}

Excel::registerType(new MoneyVndType());

// Now usable in any read or write definition!
Excel::write([
    'columns' => ['fee' => ['header' => 'Học phí', 'type' => 'money_vnd']],
    'data' => [['fee' => 1500000]],
]);
```

---

## How-To Recipes (FAQ)

<details>
<summary><strong>How to import an Excel file to an array in PHP?</strong></summary>

```php
$rows = Excel::read('file.xlsx')->rows();
```
</details>

<details>
<summary><strong>How to export a PHP array directly to Excel download?</strong></summary>

```php
Excel::write([
    'columns' => ['name' => 'Name', 'email' => 'Email'],
    'data'    => $users,
])->download('users.xlsx');
```
</details>

<details>
<summary><strong>How to export multiple sheets?</strong></summary>

```php
Excel::write([
    'sheets' => [
        ['name' => 'Users', 'columns' => [...], 'data' => $users],
        ['name' => 'Orders', 'columns' => [...], 'data' => $orders],
    ],
])->save('export.xlsx');
```
</details>

---

## License

This package is open-sourced software licensed under the [MIT license](LICENSE).
