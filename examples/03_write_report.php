<?php

declare(strict_types=1);

require dirname(__DIR__) . '/vendor/autoload.php';

use ExcelHelper\Excel;

$tempPath = sys_get_temp_dir() . '/example_03_' . uniqid() . '.xlsx';

Excel::write([
    'title' => 'BÁO CÁO DOANH SỐ BÁN HÀNG',
    'columns' => [
        'code'     => 'Mã SP',
        'name'     => ['header' => 'Tên Sản Phẩm', 'width' => 25],
        'quantity' => ['header' => 'Số Lượng', 'type' => 'int'],
        'revenue'  => ['header' => 'Doanh Thu', 'type' => 'number', 'format' => '#,##0 "₫"'],
    ],
    'data' => [
        ['code' => 'SP01', 'name' => 'Bàn phím cơ', 'quantity' => 10, 'revenue' => 15000000],
        ['code' => 'SP02', 'name' => 'Chuột không dây', 'quantity' => 25, 'revenue' => 12500000],
        ['code' => 'SP03', 'name' => 'Màn hình 27 inch', 'quantity' => 5, 'revenue' => 30000000],
    ],
    'totals' => ['revenue' => 'sum', 'quantity' => 'sum'],
    'totals_label' => 'TỔNG CỘNG',
    'freeze_header' => true,
    'zebra' => true,
    'borders' => 'thin',
])->save($tempPath);

echo "Exported styled sales report to: {$tempPath}\n";

$info = Excel::info($tempPath);
echo "Sheet: {$info[0]->name}, Total Rows: {$info[0]->totalRows}, Total Columns: {$info[0]->totalColumns}\n";

@unlink($tempPath);
exit(0);
