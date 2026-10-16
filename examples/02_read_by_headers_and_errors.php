<?php

declare(strict_types=1);

require dirname(__DIR__) . '/vendor/autoload.php';

use ExcelHelper\Excel;

$tempPath = sys_get_temp_dir() . '/example_02_' . uniqid() . '.xlsx';

Excel::write([
    'headers' => ['Mã Nhân Viên', 'Họ Và Tên', 'Tuổi', 'Email'],
    'data' => [
        ['NV01', 'Nguyễn Văn A', 30, 'a@example.com'],
        ['',     'Trần Thị B',   'invalid_age', 'b@example.com'],
    ],
])->save($tempPath);

$result = Excel::read($tempPath, [
    'header_row' => 1,
    'locale' => 'vi',
    'columns' => [
        ['header' => 'Mã Nhân Viên', 'key' => 'code', 'required' => true, 'default' => 'CHƯA_CÓ'],
        ['header' => 'Họ Và Tên',   'key' => 'name'],
        ['header' => 'Tuổi',         'key' => 'age', 'type' => 'int', 'default' => 0],
        ['header' => 'Email',        'key' => 'email', 'optional' => true],
    ],
]);

echo "Rows read: {$result->rowCount()}\n";
foreach ($result->rows() as $row) {
    echo "- Mã: {$row['code']}, Tên: {$row['name']}, Tuổi: {$row['age']}\n";
}

if ($result->hasErrors()) {
    echo "\nDanh sách lỗi dữ liệu phát hiện được:\n";
    foreach ($result->errors() as $err) {
        echo "- Dòng {$err->row}, cột {$err->column} ({$err->key}) [{$err->code}]: {$err->message}\n";
    }
}

@unlink($tempPath);
exit(0);
