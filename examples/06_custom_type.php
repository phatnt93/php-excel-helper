<?php

declare(strict_types=1);

require dirname(__DIR__) . '/vendor/autoload.php';

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Excel;
use ExcelHelper\Types\CellValue;
use ExcelHelper\Types\ColumnContext;

// 1. Define custom Money VND type
$vndType = new class implements TypeInterface {
    public function name(): string
    {
        return 'vnd';
    }

    public function read(mixed $value, ColumnContext $context): float
    {
        if (is_numeric($value)) {
            return (float) $value;
        }

        return (float) preg_replace('/[^\d]/u', '', (string) $value);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        return CellValue::numeric((float) $value, '#,##0 "₫"');
    }
};

// 2. Register with Excel helper
Excel::registerType($vndType);

$tempPath = sys_get_temp_dir() . '/example_06_' . uniqid() . '.xlsx';

Excel::write([
    'columns' => [
        'item' => 'Hạng Mục',
        'cost' => ['header' => 'Đơn Giá', 'type' => 'vnd'],
    ],
    'data' => [
        ['item' => 'Hosting Pro', 'cost' => 2500000],
        ['item' => 'Tên miền .vn', 'cost' => 750000],
    ],
    'totals' => ['cost' => 'sum'],
])->save($tempPath);

$result = Excel::read($tempPath, [
    'header_row' => 1,
    'columns' => [
        ['header' => 'Hạng Mục', 'key' => 'item'],
        ['header' => 'Đơn Giá',  'key' => 'cost', 'type' => 'vnd'],
    ],
]);

echo "Exported and read back with custom 'vnd' type:\n";
foreach ($result->rows() as $row) {
    echo "- {$row['item']}: " . number_format($row['cost'], 0, ',', '.') . " VNĐ\n";
}

@unlink($tempPath);
exit(0);
