<?php

declare(strict_types=1);

require dirname(__DIR__) . '/vendor/autoload.php';

use ExcelHelper\Excel;
use ExcelHelper\Writer\SheetBuilder;

$tempPath = sys_get_temp_dir() . '/example_04_' . uniqid() . '.xlsx';

Excel::writer()
    ->properties(['creator' => 'Company Admin', 'title' => 'Multi-Sheet Export'])
    ->default('zebra', true)
    ->sheet('Khách hàng', function (SheetBuilder $s): void {
        $s->title('DANH SÁCH KHÁCH HÀNG')
            ->column('id', 'Mã KH')
            ->column('name', 'Họ Tên', width: 25)
            ->column('vip', 'Hạng VIP', type: 'bool', labels: [true => 'VIP', false => 'Thường'])
            ->data([
                ['id' => 'KH01', 'name' => 'Công ty ABC', 'vip' => true],
                ['id' => 'KH02', 'name' => 'Nguyễn Thị C', 'vip' => false],
            ]);
    })
    ->sheet('Đơn hàng', function (SheetBuilder $s): void {
        $s->title('DANH SÁCH ĐƠN HÀNG')
            ->column('order_no', 'Mã Đơn')
            ->column('amount', 'Tổng Tiền', type: 'number', format: '#,##0')
            ->data([
                ['order_no' => 'DH-1001', 'amount' => 5400000],
                ['order_no' => 'DH-1002', 'amount' => 1200000],
            ])
            ->totals(['amount' => 'sum']);
    })
    ->save($tempPath);

echo "Exported multi-sheet workbook to: {$tempPath}\n";
$sheets = Excel::sheetNames($tempPath);
echo "Sheets inside workbook: " . implode(', ', $sheets) . "\n";

@unlink($tempPath);
exit(0);
