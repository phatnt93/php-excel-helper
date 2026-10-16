# PHP Excel Helper (Tiếng Việt)

[![Latest Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/phatnt93/php-excel-helper)
[![PHP Version](https://img.shields.io/badge/php-%5E8.1-777bb4.svg)](https://php.net)
[![PHPStan](https://img.shields.io/badge/PHPStan-level%208-brightgreen.svg)](https://phpstan.org)
[![Tests](https://img.shields.io/badge/tests-100%25%20passing-brightgreen.svg)](https://phpunit.de)
[![Coverage](https://img.shields.io/badge/coverage-%E2%89%A590%25-brightgreen.svg)](https://github.com/phatnt93/php-excel-helper)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Thư viện PHP hiện đại, an toàn và tinh gọn giúp đọc và ghi các định dạng bảng tính Excel và CSV (`.xlsx`, `.xls`, `.ods`, `.csv`) bằng mảng cấu hình PHP trực quan hoặc Fluent Builder. Được xây dựng trên nền tảng [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet).

**[Switch to English Documentation](README.md)**

---

## Mục lục

- [Tại sao nên dùng php-excel-helper?](#tại-sao-nên-dùng-php-excel-helper)
- [Cài đặt](#cài-đặt)
- [Khởi động nhanh](#khởi-động-nhanh)
  - [Đọc file thành mảng](#đọc-file-thành-mảng)
  - [Ghi mảng thành file Excel](#ghi-mảng-thành-file-excel)
  - [Fluent Builder](#fluent-builder)
- [Đọc File Chi Tiết](#đọc-file-chi-tiết)
  - [Bảng tùy chọn đọc](#bảng-tùy-chọn-đọc)
  - [3 cách khai báo cột](#3-cách-khai-báo-cột)
  - [Thuộc tính của cột đọc](#thuộc-tính-của-cột-đọc)
  - [Các kiểu dữ liệu hỗ trợ](#các-kiểu-dữ-liệu-hỗ-trợ)
  - [Xử lý lỗi dữ liệu (`on_error`)](#xử-lý-lỗi-dữ-liệu-on_error)
- [Ghi File Chi Tiết](#ghi-file-chi-tiết)
  - [Tùy chọn cấp Workbook](#tùy-chọn-cấp-workbook)
  - [Tùy chọn cấp Sheet](#tùy-chọn-cấp-sheet)
  - [Hệ thống Style thân thiện](#hệ-thống-style-thân-thiện)
  - [Các đầu ra (Lưu, Tải về, Chuỗi nhị phân)](#các-đầu-ra)
- [Xử Lý File Lớn & Tiết Kiệm Bộ Nhớ](#xử-lý-file-lớn--tiết-kiệm-bộ-nhớ)
  - [Đọc theo khối (Chunk)](#đọc-theo-khối-chunk)
  - [Đọc CSV theo luồng O(1) RAM](#đọc-csv-theo-luồng-o1-ram)
- [Tiện ích](#tiện-ích)
- [Bảo mật hàng đầu](#bảo-mật-hàng-đầu)
- [Kiểu dữ liệu tùy biến (Custom Types)](#kiểu-dữ-liệu-tùy-biến-custom-types)
- [Câu hỏi thường gặp (FAQ)](#câu-hỏi-thường-gặp-faq)
- [Giấy phép](#giấy-phép)

---

## Tại sao nên dùng php-excel-helper?

[PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) là thư viện đọc/ghi bảng tính chuẩn mực của hệ sinh thái PHP, nhưng API ở mức rất thấp khiến các tác vụ quen thuộc tốn hàng chục dòng mã: phải tự duyệt từng ô, tự đổi số serial thành ngày, tự giữ số 0 đầu cho số điện thoại, tự viết mảng style lồng 5 cấp, và tự cấu hình bộ lọc đọc khối.

`php-excel-helper` là **lớp nghiệp vụ tinh giản** đặt phía trên PhpSpreadsheet:

| Tính năng | PhpSpreadsheet thuần | php-excel-helper |
|---|---|---|
| **Phong cách API** | Hướng đối tượng cấp thấp, duyệt ô thủ công | Mảng cấu hình thuần túy hoặc Fluent Builder song song |
| **Ngày tháng Excel** | Trả về số thực serial (vd: `46311.5`), cần tự đổi | Tự động chuyển thành chuỗi `Y-m-d` hoặc `DateTimeImmutable` |
| **Số điện thoại / Số 0 đầu** | Bị ép thành số nguyên làm mất số 0 (`901234567`) | Giữ nguyên chuỗi text an toàn (`'0901234567'`) |
| **Xử lý lỗi dòng dữ liệu** | Phải tự viết logic kiểm tra từng ô | Báo lỗi chi tiết theo dòng (`required`, `invalid_int`, `custom`...) |
| **Chống Formula Injection** | Lập trình viên phải tự escape | Tự động khử độc (`TYPE_STRING` trong Excel, escape trong CSV) |
| **Định dạng giao diện (Style)** | Mảng lồng nhiều cấp (`applyFromArray`) | Cấu hình phẳng thân thiện (`'bold' => true, 'bg' => '#1F4E78'`) |
| **File CSV khổng lồ** | Tải toàn bộ vào RAM (dễ tràn bộ nhớ) | Đọc tuần tự bằng `SplFileObject` với bộ nhớ O(1) không đổi |
| **Lối thoát linh hoạt** | Khó can thiệp sâu khi thư viện bọc quá kín | Luôn có lối thoát trực tiếp qua `->spreadsheet()` hoặc `tap($sheet)` |

---

## Cài đặt

```bash
composer require phatnt93/php-excel-helper
```

**Yêu cầu hệ thống:**
- PHP 8.1 trở lên
- Extension: `mbstring` (có sẵn mặc định trên hầu hết hosting/server)
- Khuyến nghị: `ext-intl` (để chuẩn hóa Unicode NFC khi so khớp tiêu đề tiếng Việt)

---

## Khởi động nhanh

### Đọc file thành mảng

```php
use ExcelHelper\Excel;

$result = Excel::read('nhan-vien.xlsx', [
    'header_row' => 1,
    'columns' => [
        ['header' => 'Mã NV',      'key' => 'code'],
        ['header' => 'Họ và tên',  'key' => 'name'],
        ['header' => 'Ngày sinh',  'key' => 'birthday', 'type' => 'date'],
        ['header' => 'Lương',      'key' => 'salary',   'type' => 'number'],
        ['header' => 'Trạng thái', 'key' => 'active',   'type' => 'bool'],
    ],
]);

// Trả về mảng kết hợp gọn gàng
foreach ($result->rows() as $row) {
    echo "{$row['code']}: {$row['name']} - Lương: {$row['salary']}\n";
}

// Kiểm tra lỗi dữ liệu nếu có
if ($result->hasErrors()) {
    foreach ($result->errors() as $error) {
        echo "Dòng {$error->row}, cột {$error->column} ({$error->key}): {$error->message}\n";
    }
}
```

### Ghi mảng thành file Excel

```php
use ExcelHelper\Excel;

Excel::write([
    'title' => 'BẢNG LƯƠNG NHÂN VIÊN THÁNG 10/2026',
    'columns' => [
        'code'   => 'Mã NV',
        'name'   => ['header' => 'Họ và tên', 'width' => 28],
        'salary' => ['header' => 'Mức lương', 'type' => 'number', 'format' => '#,##0'],
        'active' => ['header' => 'Đang làm', 'type' => 'bool', 'labels' => [true => 'Đang làm', false => 'Nghỉ việc']],
    ],
    'data' => [
        ['code' => 'NV001', 'name' => 'Nguyễn Văn A', 'salary' => 15000000, 'active' => true],
        ['code' => 'NV002', 'name' => 'Trần Thị B',   'salary' => 18500000, 'active' => false],
    ],
    'totals' => ['salary' => 'sum'],
    'freeze_header' => true,
    'zebra' => true,
])->save('bang-luong.xlsx');
```

### Fluent Builder

```php
// Đọc bằng builder
$result = Excel::reader('nhan-vien.xlsx')
    ->headerRow(1)
    ->column('A', 'code')
    ->column('B', 'name')
    ->column('C', 'salary', type: 'number')
    ->onError('skip_row')
    ->read();

// Ghi bằng builder
Excel::writer()
    ->title('Danh Sách')
    ->column('code', 'Mã')
    ->column('name', 'Tên', width: 25)
    ->data($data)
    ->save('xuat.xlsx');
```

---

## Đọc File Chi Tiết

### Bảng tùy chọn đọc

| Tùy chọn | Kiểu | Mặc định | Ý nghĩa |
|---|---|---|---|
| `format` | `string` | `'auto'` | Định dạng file (`'auto'`, `'xlsx'`, `'xls'`, `'ods'`, `'csv'`). Tự nhận diện theo nội dung. |
| `allowed_formats` | `list<string>` | `['xlsx','xls','ods','csv']` | Danh sách trắng các định dạng cho phép. |
| `sheet` | `int\|string` | `0` | Chỉ số sheet (0-based) hoặc tên sheet. |
| `header_row` | `?int` | `null` | Số dòng tiêu đề (1-based). Cho phép ánh xạ cột theo tên tiêu đề. |
| `start_row` | `int` | `header_row + 1` hoặc `1` | Dòng dữ liệu bắt đầu (1-based). |
| `end_row` | `?int` | `null` | Dòng dữ liệu kết thúc (`null` = đọc đến hết dữ liệu). |
| `columns` | `array` | `[]` | Khai báo các cột cần đọc. Rỗng = đọc toàn bộ dạng chuỗi thô. |
| `skip_empty_rows` | `bool` | `true` | Tự động bỏ qua các dòng hoàn toàn trắng. |
| `stop_at_empty_row`| `bool` | `false` | Dừng đọc ngay khi gặp dòng trắng đầu tiên. |
| `trim` | `bool` | `true` | Cắt khoảng trắng đầu/cuối của chuỗi. |
| `on_error` | `string` | `'collect'` | Chế độ xử lý lỗi: `'collect'`, `'skip_row'`, hoặc `'throw'`. |
| `row_number_key` | `?string` | `null` | Đặt tên key (vd: `'_row'`) để lưu số dòng Excel gốc vào kết quả. |
| `calculate_formulas`| `bool` | `false` | `false` = lấy giá trị đã tính sẵn trong file; `true` = chạy bộ tính toán. |
| `max_rows` | `?int` | `null` | Giới hạn tối đa số dòng cho phép (chống DoS file khổng lồ). |
| `max_file_size` | `int\|string\|null`| `null` | Giới hạn dung lượng file (`'20M'`, `'512K'`, hoặc số byte). |
| `date_format` | `?string` | `'Y-m-d'` | Định dạng chuỗi ngày ra (`null` = trả về `DateTimeImmutable`). |
| `datetime_format` | `?string` | `'Y-m-d H:i:s'` | Định dạng ngày giờ ra. |
| `csv` | `array` | `[]` | Tùy chọn CSV: `delimiter` (mặc định dò tự động), `enclosure`, `input_encoding`. |
| `locale` | `?string` | `null` | Ngôn ngữ thông báo lỗi dữ liệu (`'en'`, `'vi'`). |

---

### 3 cách khai báo cột

#### (a) Theo chữ cột hoặc số cột (key mảng)
```php
'columns' => [
    'A' => 'code',             // Rút gọn của ['key' => 'code', 'type' => 'string']
    'B' => ['key' => 'name'],
    3   => ['key' => 'salary', 'type' => 'float'], // Cột số 3 = C
]
```

#### (b) Danh sách tuần tự từ `start_column`
```php
'start_column' => 'B',
'columns' => [
    'code',                     // B
    'name',                     // C
    null,                       // D bị bỏ qua
    ['key' => 'salary', 'type' => 'float'], // E
]
```

#### (c) Theo tên tiêu đề (cần `header_row`)
```php
'header_row' => 1,
'columns' => [
    ['header' => 'Mã Nhân Viên', 'key' => 'code'],
    ['header' => ['Email', 'Hòm thư điện tử'], 'key' => 'email', 'optional' => true],
]
```

---

### Thuộc tính của cột đọc

| Thuộc tính | Kiểu | Mặc định | Ý nghĩa |
|---|---|---|---|
| `key` | `string` | — | Bắt buộc. Tên key trong mảng kết quả trả về. |
| `type` | `string` | `'string'` | Kiểu dữ liệu ép (`'string'`, `'int'`, `'float'`, `'decimal'`, `'bool'`, `'date'`, `'datetime'`, `'time'`, `'raw'`, `'formatted'`). |
| `format` | `?string` | theo kiểu | Định dạng ngày ra (vd: `'d/m/Y'`) hoặc số chữ số thập phân cho decimal. |
| `input_format` | `string\|list<string>` | `null` | Định dạng để parse khi ô là chuỗi văn bản (vd: `'d/m/Y'`). |
| `default` | `mixed` | `null` | Giá trị mặc định khi ô rỗng hoặc khi có lỗi dữ liệu. |
| `required` | `bool` | `false` | Bắt buộc ô không được rỗng. |
| `trim` | `?bool` | theo global | Ghi đè cấu hình trim cho riêng cột này. |
| `map` | `?array` | `null` | Bảng ánh xạ giá trị trước khi ép kiểu (vd: `['Nam' => 'male', 'Nữ' => 'female']`). |
| `map_strict` | `bool` | `false` | Báo lỗi `not_in_map` nếu giá trị không có trong bảng map. |
| `transform` | `?callable` | `null` | Hàm biến đổi sau khi ép kiểu: `fn($val, RowContext $ctx): mixed`. |
| `validate` | `?callable` | `null` | Hàm kiểm tra hợp lệ: `fn($val, RowContext $ctx): true|string`. |
| `optional` | `bool` | `false` | Cho phép cột tiêu đề vắng mặt mà không báo lỗi. |

---

### Xử lý lỗi dữ liệu (`on_error`)

- **`collect` (mặc định)**: Ghi nhận lỗi vào danh sách `$result->errors()`, gán giá trị `default`, và **vẫn giữ** dòng đó trong `$result->rows()`.
- **`skip_row`**: Ghi nhận lỗi và **loại bỏ** dòng có lỗi khỏi `$result->rows()`.
- **`throw`**: Ném ngay ngoại lệ `RowValidationException` ở lỗi đầu tiên bắt gặp.

---

## Ghi File Chi Tiết

### Hệ thống Style thân thiện

Không cần viết mảng lồng phức tạp, bạn có thể truyền các thuộc tính phẳng:

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
        'code'   => ['header' => 'Mã', 'align' => 'center'],
        'amount' => ['header' => 'Số tiền', 'style' => ['bold' => true]],
    ],
    'borders' => 'thin',
    'zebra'   => true,
    'data'    => $rows,
])->save('bao-cao.xlsx');
```

### Các đầu ra

```php
$out = Excel::write($config);

// 1. Lưu file trên máy chủ
$out->save('/duong-dan/file.xlsx');

// 2. Tải về trình duyệt trực tiếp kèm header UTF-8 an toàn
$out->download('bao-cao.xlsx');

// 3. Lấy chuỗi nhị phân (binary string)
$binary = $out->toString();

// 4. Lối thoát sang PhpSpreadsheet thuần
$spreadsheet = $out->spreadsheet();
```

---

## Xử Lý File Lớn & Tiết Kiệm Bộ Nhớ

### Đọc theo khối (Chunk)

Đọc file Excel hàng trăm ngàn dòng mà không gây tràn RAM:

```php
$summary = Excel::chunk('big-data.xlsx', ['columns' => [...]], 1000,
    function (array $chunkRows, ChunkContext $ctx): void {
        echo "Đang xử lý khối #{$ctx->chunkIndex()} ({$ctx->rowCount} dòng)...\n";

        // Lưu vào cơ sở dữ liệu
        DB::table('users')->insert($chunkRows);

        // Có thể dừng sớm nếu cần:
        // if ($condition) { return false; }
    }
);

echo "Tổng số dòng đã xử lý: {$summary->totalRows()}\n";
```

### Đọc CSV theo luồng O(1) RAM

Khi đọc file CSV ở chế độ chunk, `php-excel-helper` tự động chuyển sang đọc luồng qua `SplFileObject`, đảm bảo **bộ nhớ RAM ổn định tuyệt đối dưới 10MB** ngay cả với file CSV dung lượng hàng gigabyte.

---

## Tiện ích

```php
// Liệt kê danh sách sheet mà không cần load ô nào
$sheets = Excel::sheetNames('so-tay.xlsx');

// Lấy thông tin kích thước của các sheet
$info = Excel::info('so-tay.xlsx');
foreach ($info as $s) {
    echo "Sheet {$s->name}: {$s->totalRows} dòng, {$s->totalColumns} cột\n";
}

// Đọc nhanh dòng tiêu đề
$headers = Excel::headers('so-tay.xlsx', row: 1);
```

---

## Điền Dữ Liệu Vào File Mẫu (`Excel::template`)

Điền dữ liệu vào các tệp Excel mẫu có sẵn (hóa đơn, phiếu thu, hợp đồng) mà vẫn giữ nguyên logo, phông chữ, định dạng ô và các công thức phức tạp:

```php
Excel::template('mau_hoa_don.xlsx')
    ->set('customer_name', 'Công ty Cổ phần Công nghệ ABC')
    ->set('invoice_no', 'HD-2026-001')
    ->setValues([
        'ngay_lap' => '16/10/2026',
        'ghi_chu' => 'Đã thanh toán chuyển khoản',
        'B15' => 'Ghi trực tiếp vào tọa độ ô',
    ])
    ->insertTable('item_name', [
        ['item_name' => 'Gói Hosting Doanh nghiệp', 'qty' => 1, 'price' => 2500000],
        ['item_name' => 'Tên miền .com.vn', 'qty' => 2, 'price' => 750000],
    ])
    ->save('hoa_don_001.xlsx');
```

* `insertTable()` tự động kế thừa định dạng ô và tự động đẩy các dòng chân bảng / dòng tổng phía dưới xuống!

---

## Các Tính Năng Nâng Cao v1.1.0

### 1. Đọc Theo Luồng O(1) RAM (`Excel::stream`)

Đọc tệp XLSX và CSV siêu lớn theo từng dòng bằng PHP `Generator` với mức tiêu hao bộ nhớ không đổi:

```php
foreach (Excel::stream('file_khong_lo.xlsx', ['start_row' => 2]) as $rowNumber => $row) {
    // Xử lý từng dòng mượt mà trên server RAM thấp
}
```

### 2. Chèn Ảnh & Logo

Chèn logo công ty, ảnh sản phẩm với tọa độ ô và kích thước pixel:

```php
Excel::write([
    'images' => [
        ['path' => 'assets/logo.png', 'cell' => 'A1', 'width' => 120, 'height' => 50],
    ],
    'columns' => [...],
    'data' => [...],
])->save('bao_cao.xlsx');
```

### 3. Ràng Buộc Dữ Liệu & Dropdown List (Data Validation)

Tạo danh sách chọn dropdown và ràng buộc dữ liệu trực tiếp trong Excel:

```php
'status' => [
    'header' => 'Trạng thái',
    'validation' => ['Chờ duyệt', 'Đã duyệt', 'Từ chối'], // Dropdown list
],
'age' => [
    'header' => 'Tuổi',
    'validation' => ['type' => 'integer', 'min' => 18, 'max' => 65],
],
```

### 4. Định Dạng Có Điều Kiện (Conditional Formatting)

Tự động đổi màu ô dựa trên giá trị:

```php
'balance' => [
    'header' => 'Số dư',
    'conditional' => [
        ['operator' => '<', 'value' => 0, 'color' => 'FF0000', 'bg' => 'FFE0E0', 'bold' => true],
    ],
],
```

---

## Giấy phép

Mã nguồn được phát hành theo [Giấy phép MIT](LICENSE).
