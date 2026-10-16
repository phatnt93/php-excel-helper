<?php

declare(strict_types=1);

require_once dirname(__DIR__) . '/vendor/autoload.php';

use ExcelHelper\Excel;

echo "=== Demonstration of Advanced v1.1.0 Features ===" . PHP_EOL . PHP_EOL;

// 1. Image Insertion, Data Validation, and Conditional Formatting
$exportPath = sys_get_temp_dir() . '/example_07_advanced_' . uniqid() . '.xlsx';

// Generate tiny sample logo
$logoPath = sys_get_temp_dir() . '/sample_logo_' . uniqid() . '.png';
$img = imagecreatetruecolor(50, 50);
if ($img !== false) {
    $bg = imagecolorallocate($img, 41, 128, 185);
    if ($bg !== false) {
        imagefill($img, 0, 0, $bg);
    }
    imagepng($img, $logoPath);
    imagedestroy($img);
}

Excel::write([
    'title' => 'BÁO CÁO TÀI CHÍNH & NHÂN SỰ',
    'images' => [
        ['path' => $logoPath, 'cell' => 'D1', 'width' => 40, 'height' => 40],
    ],
    'conditional_formatting' => [
        [
            'range' => 'C3:C10',
            'operator' => '<',
            'value' => 0,
            'color' => 'FF0000',
            'bg' => 'FFE0E0',
            'bold' => true,
        ],
    ],
    'columns' => [
        'name' => 'Họ và tên',
        'status' => [
            'header' => 'Trạng thái',
            'validation' => ['Chờ duyệt', 'Đã duyệt', 'Từ chối'],
        ],
        'balance' => [
            'header' => 'Số dư tài khoản',
            'type' => 'number',
            'format' => '#,##0',
        ],
    ],
    'data' => [
        ['name' => 'Nguyễn Văn A', 'status' => 'Đã duyệt', 'balance' => 15000000],
        ['name' => 'Trần Thị B', 'status' => 'Chờ duyệt', 'balance' => -2500000],
        ['name' => 'Lê Văn C', 'status' => 'Đã duyệt', 'balance' => 8400000],
    ],
])->save($exportPath);

echo "1. Exported workbook with Logo, Dropdown Validation, and Conditional Formatting to: {$exportPath}" . PHP_EOL;

// 2. Fast Streaming XLSX Reader
echo PHP_EOL . "2. Streaming rows one-by-one with O(1) memory via Excel::stream():" . PHP_EOL;
foreach (Excel::stream($exportPath, ['header_row' => 3, 'start_row' => 4]) as $rowNum => $row) {
    printf("   Line %d: %-15s | Status: %-10s | Balance: %10s\n", $rowNum, $row['Họ và tên'], $row['Trạng thái'], $row['Số dư tài khoản']);
}

// 3. Template Engine
echo PHP_EOL . "3. Processing template with Excel::template():" . PHP_EOL;
$templateFile = sys_get_temp_dir() . '/invoice_tpl_' . uniqid() . '.xlsx';

Excel::write([
    'show_header' => false,
    'data' => [
        ['HÓA ĐƠN BÁN HÀNG'],
        ['Khách hàng: {{customer_name}}', 'Mã HĐ: {{invoice_no}}'],
        ['Sản phẩm', 'Số lượng', 'Đơn giá'],
        ['{{item_name}}', '{{qty}}', '{{price}}'],
        ['Cảm ơn quý khách đã mua hàng!'],
    ],
])->save($templateFile);

$filledInvoice = sys_get_temp_dir() . '/invoice_filled_' . uniqid() . '.xlsx';
Excel::template($templateFile)
    ->set('customer_name', 'Công ty TNHH Giải Pháp Phần Mềm')
    ->set('invoice_no', 'HD-2026-99')
    ->insertTable('item_name', [
        ['item_name' => 'Gói Hosting Pro', 'qty' => 1, 'price' => 2500000],
        ['item_name' => 'Tên miền .com.vn', 'qty' => 2, 'price' => 750000],
        ['item_name' => 'Chứng chỉ SSL EV', 'qty' => 1, 'price' => 4500000],
    ])
    ->save($filledInvoice);

echo "   Successfully filled template and shifted footer rows down to: {$filledInvoice}" . PHP_EOL;

@unlink($logoPath);
@unlink($exportPath);
@unlink($templateFile);
@unlink($filledInvoice);

echo PHP_EOL . "=== All 5 advanced features executed successfully! ===" . PHP_EOL;
