<?php

declare(strict_types=1);

require dirname(__DIR__) . '/vendor/autoload.php';

use ExcelHelper\Excel;

// 1. Generate sample spreadsheet file
$tempPath = sys_get_temp_dir() . '/example_01_' . uniqid() . '.xlsx';

Excel::write([
    'columns' => [
        'id' => 'ID',
        'name' => 'Name',
        'price' => ['header' => 'Price', 'type' => 'number'],
        'created_at' => ['header' => 'Created At', 'type' => 'date'],
    ],
    'data' => [
        ['id' => 'P001', 'name' => 'Keyboard', 'price' => 45.5, 'created_at' => '2026-10-16'],
        ['id' => 'P002', 'name' => 'Mouse',    'price' => 20.0, 'created_at' => '2026-10-15'],
    ],
])->save($tempPath);

// 2. Read back using array configuration
$result = Excel::read($tempPath, [
    'header_row' => 1,
    'columns' => [
        'A' => 'id',
        'B' => 'name',
        'C' => ['key' => 'price', 'type' => 'float'],
        'D' => ['key' => 'date', 'type' => 'date'],
    ],
]);

echo "Read {$result->rowCount()} products:\n";
foreach ($result->rows() as $row) {
    echo "- [{$row['id']}] {$row['name']}: \${$row['price']} (Date: {$row['date']})\n";
}

@unlink($tempPath);
exit(0);
