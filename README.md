# PHP Excel Helper

> **Status: under development — not released yet.** The API below is the planned design and may change before `v1.0.0`.

Read and write Excel/spreadsheet files (xlsx, xls, ods, csv) with plain PHP arrays. Built on top of [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet).

```php
use ExcelHelper\Excel;

// Read: typed columns, defaults and row-level errors
$result = Excel::read('employees.xlsx', [
    'header_row' => 1,
    'start_row'  => 2,
    'columns'    => [
        'A' => ['key' => 'code', 'required' => true],
        'B' => 'name',
        'C' => ['key' => 'birthday', 'type' => 'date', 'input_format' => 'd/m/Y'],
        'D' => ['key' => 'salary', 'type' => 'float', 'default' => 0],
    ],
]);

// Write: headers, typed cells, number formats, totals
Excel::write([
    'name'    => 'Employees',
    'columns' => [
        'code'   => 'Code',
        'name'   => ['header' => 'Full name', 'width' => 28],
        'salary' => ['header' => 'Salary', 'type' => 'number', 'format' => '#,##0'],
    ],
    'data'    => $employees,
    'totals'  => ['salary' => 'sum'],
])->save('employees.xlsx');
```

## Requirements

- PHP 8.1+
- `phpoffice/phpspreadsheet` ^5.8 (installed automatically)

## License

MIT — see [LICENSE](LICENSE).
