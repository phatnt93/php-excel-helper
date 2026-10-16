<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Integration;

use ExcelHelper\Excel;
use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\RowValidationException;
use ExcelHelper\Exceptions\SheetNotFoundException;
use ExcelHelper\Reader\RowContext;
use ExcelHelper\Tests\TestCase;

final class ReaderIntegrationTest extends TestCase
{
    public function testReadByColumnLetters(): void
    {
        $path = $this->createSpreadsheetFile([
            ['NV001', 'Nguyen Van A', 15000000],
            ['NV002', 'Tran Thi B', 18000000],
        ], 'xlsx');

        $result = Excel::read($path, [
            'columns' => [
                'A' => 'code',
                'B' => 'name',
                'C' => ['key' => 'salary', 'type' => 'float'],
            ],
        ]);

        $this->assertSame(2, $result->rowCount());
        $this->assertFalse($result->hasErrors());

        $rows = $result->rows();
        $this->assertSame('NV001', $rows[0]['code']);
        $this->assertSame('Nguyen Van A', $rows[0]['name']);
        $this->assertSame(15000000.0, $rows[0]['salary']);
    }

    public function testReadBySequentialListWithNullSkip(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Header_A', 'NV001', 'Ignored_C', 25],
            ['Header_A', 'NV002', 'Ignored_C', 30],
        ], 'xlsx');

        $result = Excel::read($path, [
            'start_column' => 'B',
            'columns' => [
                'code',
                null, // Skip column C
                ['key' => 'age', 'type' => 'int'],
            ],
        ]);

        $this->assertSame(2, $result->rowCount());
        $rows = $result->rows();
        $this->assertSame('NV001', $rows[0]['code']);
        $this->assertSame(25, $rows[0]['age']);
        $this->assertArrayNotHasKey('Ignored_C', $rows[0]);
    }

    public function testReadByHeaderRow(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Mã Nhân Viên', 'Họ Và Tên', 'Email'],
            ['NV001', 'Nguyen Van A', 'a@example.com'],
            ['NV002', 'Tran Thi B', 'b@example.com'],
        ], 'xlsx');

        $result = Excel::read($path, [
            'header_row' => 1,
            'columns' => [
                ['header' => 'Mã Nhân Viên', 'key' => 'code'],
                ['header' => 'Họ Và Tên', 'key' => 'name'],
                ['header' => 'Email', 'key' => 'email'],
            ],
        ]);

        $this->assertSame(2, $result->rowCount());
        $this->assertSame('NV001', $result->rows()[0]['code']);
        $this->assertSame('b@example.com', $result->rows()[1]['email']);
        $this->assertSame('Mã Nhân Viên', $result->headers()['A']);
    }

    public function testReadDifferentSpreadsheetFormats(): void
    {
        $rows = [
            ['Code', 'Price'],
            ['SKU-1', 99.5],
            ['SKU-2', 150.0],
        ];

        $formats = ['xlsx', 'ods', 'csv'];

        foreach ($formats as $fmt) {
            $path = $this->createSpreadsheetFile($rows, $fmt);

            $result = Excel::read($path, [
                'header_row' => 1,
                'columns' => [
                    ['header' => 'Code', 'key' => 'code'],
                    ['header' => 'Price', 'key' => 'price', 'type' => 'float'],
                ],
            ]);

            $this->assertSame(2, $result->rowCount(), "Failed for format {$fmt}");
            $this->assertSame('SKU-1', $result->rows()[0]['code'], "Failed code for {$fmt}");
            $this->assertSame(99.5, $result->rows()[0]['price'], "Failed price for {$fmt}");
        }
    }

    public function testReadDateSystem1900and1904(): void
    {
        // 1900 calendar file with serial number 44000
        $path1900 = $this->createSpreadsheetFile([
            ['Birthday'],
            [44000],
        ], 'xlsx');

        // Excel::read on 1900
        $res1900 = Excel::read($path1900, [
            'header_row' => 1,
            'columns' => [
                ['header' => 'Birthday', 'key' => 'birthday', 'type' => 'date'],
            ],
        ]);
        $this->assertSame('2020-06-18', $res1900->rows()[0]['birthday']);

        // Excel::stream on 1900
        $stream1900 = iterator_to_array(Excel::stream($path1900, [
            'header_row' => 1,
            'start_row' => 2,
            'columns' => [
                ['header' => 'Birthday', 'key' => 'birthday', 'type' => 'date'],
            ],
        ]));
        $this->assertSame('2020-06-18', array_values($stream1900)[0]['birthday']);

        // Excel::chunk on 1900
        $chunk1900 = null;
        Excel::chunk($path1900, [
            'header_row' => 1,
            'start_row' => 2,
            'columns' => [
                ['header' => 'Birthday', 'key' => 'birthday', 'type' => 'date'],
            ],
        ], 10, function (array $rows) use (&$chunk1900): void {
            $chunk1900 = $rows[0]['birthday'];
        });
        $this->assertSame('2020-06-18', $chunk1900);

        // 1904 calendar file with serial number 44000
        $path1904 = $this->createSpreadsheetFile([
            ['Birthday'],
            [44000],
        ], 'xlsx', 'Sheet1', ['date_system_1904' => true]);

        // Excel::read on 1904
        $res1904 = Excel::read($path1904, [
            'header_row' => 1,
            'columns' => [
                ['header' => 'Birthday', 'key' => 'birthday', 'type' => 'date'],
            ],
        ]);
        $this->assertSame('2024-06-19', $res1904->rows()[0]['birthday']);

        // Excel::stream on 1904
        $stream1904 = iterator_to_array(Excel::stream($path1904, [
            'header_row' => 1,
            'start_row' => 2,
            'columns' => [
                ['header' => 'Birthday', 'key' => 'birthday', 'type' => 'date'],
            ],
        ]));
        $this->assertSame('2024-06-19', array_values($stream1904)[0]['birthday']);

        // Excel::chunk on 1904
        $chunk1904 = null;
        Excel::chunk($path1904, [
            'header_row' => 1,
            'start_row' => 2,
            'columns' => [
                ['header' => 'Birthday', 'key' => 'birthday', 'type' => 'date'],
            ],
        ], 10, function (array $rows) use (&$chunk1904): void {
            $chunk1904 = $rows[0]['birthday'];
        });
        $this->assertSame('2024-06-19', $chunk1904);

        // Reset global calendar state
        \PhpOffice\PhpSpreadsheet\Shared\Date::setExcelCalendar(\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_WINDOWS_1900);
    }

    public function testReadStringPreservesLeadingZeros(): void
    {
        $path = $this->createSpreadsheetFile([
            ['0901234567'],
        ], 'xlsx');

        $result = Excel::read($path, [
            'columns' => [
                'A' => ['key' => 'phone', 'type' => 'string'],
            ],
        ]);

        $this->assertSame('0901234567', $result->rows()[0]['phone']);
    }

    public function testDefaultValueAndRequiredError(): void
    {
        $path = $this->createSpreadsheetFile([
            ['NV001', ''],
            ['', 'Present'],
        ], 'xlsx');

        $result = Excel::read($path, [
            'columns' => [
                'A' => ['key' => 'code', 'required' => true, 'default' => 'DEFAULT_CODE'],
                'B' => ['key' => 'note', 'default' => 'N/A'],
            ],
        ]);

        $this->assertTrue($result->hasErrors());
        $this->assertCount(1, $result->errors());
        $this->assertSame('required', $result->errors()[0]->code);
        $this->assertSame(2, $result->errors()[0]->row);

        $rows = $result->rows();
        $this->assertSame('NV001', $rows[0]['code']);
        $this->assertSame('N/A', $rows[0]['note']);
        $this->assertSame('DEFAULT_CODE', $rows[1]['code']);
    }

    public function testMapAndMapStrict(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Nam'],
            ['Nữ'],
            ['Khác'],
        ], 'xlsx');

        $result = Excel::read($path, [
            'columns' => [
                'A' => [
                    'key' => 'gender',
                    'map' => ['Nam' => 'male', 'Nữ' => 'female'],
                    'map_strict' => true,
                    'default' => 'unknown',
                ],
            ],
        ]);

        $this->assertTrue($result->hasErrors());
        $this->assertCount(1, $result->errors());
        $this->assertSame('not_in_map', $result->errors()[0]->code);

        $rows = $result->rows();
        $this->assertSame('male', $rows[0]['gender']);
        $this->assertSame('female', $rows[1]['gender']);
        $this->assertSame('unknown', $rows[2]['gender']);
    }

    public function testTransformAndValidateCallbacks(): void
    {
        $path = $this->createSpreadsheetFile([
            ['hello', 25],
            ['world', -5],
        ], 'xlsx');

        $result = Excel::read($path, [
            'columns' => [
                'A' => [
                    'key' => 'title',
                    'transform' => static fn(string $v, RowContext $ctx): string => strtoupper($v),
                ],
                'B' => [
                    'key' => 'age',
                    'type' => 'int',
                    'validate' => static fn(int $v, RowContext $ctx): true|string => $v >= 0 ? true : 'Age cannot be negative',
                    'default' => 0,
                ],
            ],
        ]);

        $this->assertTrue($result->hasErrors());
        $this->assertCount(1, $result->errors());
        $this->assertSame('custom', $result->errors()[0]->code);
        $this->assertStringContainsString('Age cannot be negative', $result->errors()[0]->message);

        $rows = $result->rows();
        $this->assertSame('HELLO', $rows[0]['title']);
        $this->assertSame(25, $rows[0]['age']);
        $this->assertSame('WORLD', $rows[1]['title']);
        $this->assertSame(0, $rows[1]['age']);
    }

    public function testRowNumberKey(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Header'],
            ['First'],
            ['Second'],
        ], 'xlsx');

        $result = Excel::read($path, [
            'header_row' => 1,
            'row_number_key' => '_row',
            'columns' => [
                ['header' => 'Header', 'key' => 'text'],
            ],
        ]);

        $this->assertSame(2, $result->rows()[0]['_row']);
        $this->assertSame(3, $result->rows()[1]['_row']);
    }

    public function testOnErrorModes(): void
    {
        $path = $this->createSpreadsheetFile([
            ['10'],
            ['invalid_int'],
            ['30'],
        ], 'xlsx');

        // Mode 1: collect (keeps all rows)
        $resCollect = Excel::read($path, [
            'on_error' => 'collect',
            'columns' => ['A' => ['key' => 'num', 'type' => 'int', 'default' => 0]],
        ]);
        $this->assertCount(3, $resCollect->rows());
        $this->assertCount(1, $resCollect->errors());
        $this->assertSame(0, $resCollect->rows()[1]['num']);

        // Mode 2: skip_row (drops the erroneous row)
        $resSkip = Excel::read($path, [
            'on_error' => 'skip_row',
            'columns' => ['A' => ['key' => 'num', 'type' => 'int', 'default' => 0]],
        ]);
        $this->assertCount(2, $resSkip->rows());
        $this->assertSame(10, $resSkip->rows()[0]['num']);
        $this->assertSame(30, $resSkip->rows()[1]['num']);
        $this->assertCount(1, $resSkip->errors());

        // Mode 3: throw (throws RowValidationException immediately)
        $this->expectException(RowValidationException::class);
        Excel::read($path, [
            'on_error' => 'throw',
            'columns' => ['A' => ['key' => 'num', 'type' => 'int']],
        ]);
    }

    public function testSkipEmptyRowsAndStopAtEmptyRow(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Row 1'],
            [''],
            ['Row 3'],
        ], 'xlsx');

        // skip_empty_rows = true (default)
        $resSkip = Excel::read($path, [
            'skip_empty_rows' => true,
            'columns' => ['A' => 'name'],
        ]);
        $this->assertCount(2, $resSkip->rows());
        $this->assertSame('Row 1', $resSkip->rows()[0]['name']);
        $this->assertSame('Row 3', $resSkip->rows()[1]['name']);

        // stop_at_empty_row = true
        $resStop = Excel::read($path, [
            'stop_at_empty_row' => true,
            'columns' => ['A' => 'name'],
        ]);
        $this->assertCount(1, $resStop->rows());
        $this->assertSame('Row 1', $resStop->rows()[0]['name']);
    }

    public function testMultiSheetSelection(): void
    {
        $path = $this->createMultiSheetFile([
            ['name' => 'Customers', 'rows' => [['Cust 1'], ['Cust 2']]],
            ['name' => 'Orders', 'rows' => [['Order 100'], ['Order 101']]],
        ]);

        // By index 1
        $resOrders = Excel::read($path, [
            'sheet' => 1,
            'columns' => ['A' => 'item'],
        ]);
        $this->assertSame('Orders', $resOrders->sheetName());
        $this->assertSame('Order 100', $resOrders->rows()[0]['item']);

        // By name 'Customers'
        $resCust = Excel::read($path, [
            'sheet' => 'Customers',
            'columns' => ['A' => 'item'],
        ]);
        $this->assertSame('Customers', $resCust->sheetName());
        $this->assertSame('Cust 1', $resCust->rows()[0]['item']);
    }

    public function testSheetNotFoundThrows(): void
    {
        $path = $this->createSpreadsheetFile([['Header']]);

        $this->expectException(SheetNotFoundException::class);
        Excel::read($path, ['sheet' => 'NonExistentSheet']);
    }

    public function testMissingRequiredHeaderThrows(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Header_A', 'Header_B'],
            ['Val_A', 'Val_B'],
        ]);

        $this->expectException(ReaderException::class);
        $this->expectExceptionMessage('Required header column "Header_C"');

        Excel::read($path, [
            'header_row' => 1,
            'columns' => [
                ['header' => 'Header_C', 'key' => 'c'],
            ],
        ]);
    }

    public function testOptionalHeaderColumnDefaults(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Header_A'],
            ['Val_A'],
        ]);

        $result = Excel::read($path, [
            'header_row' => 1,
            'columns' => [
                ['header' => 'Header_A', 'key' => 'a'],
                ['header' => 'Optional_Header', 'key' => 'b', 'optional' => true, 'default' => 'DEF'],
            ],
        ]);

        $this->assertSame('DEF', $result->rows()[0]['b']);
    }

    public function testFormattedTypeReading(): void
    {
        $path = $this->createSpreadsheetFile([
            ['12.5%'],
        ]);

        $result = Excel::read($path, [
            'columns' => [
                'A' => ['key' => 'percent', 'type' => 'formatted'],
            ],
        ]);

        $this->assertSame('12.5%', $result->rows()[0]['percent']);
    }

    public function testCalculateFormulasOption(): void
    {
        $path = $this->createSpreadsheetFile([
            [10, 20, '=A1+B1'],
        ]);

        $result = Excel::read($path, [
            'calculate_formulas' => true,
            'columns' => [
                'A' => ['key' => 'a', 'type' => 'int'],
                'B' => ['key' => 'b', 'type' => 'int'],
                'C' => ['key' => 'sum', 'type' => 'int'],
            ],
        ]);

        $this->assertSame(30, $result->rows()[0]['sum']);
    }

    public function testMemoryStabilityOnLoopReads(): void
    {
        $path = $this->createSpreadsheetFile([
            ['NV001', 'Nguyen Van A', 1000],
            ['NV002', 'Tran Thi B', 2000],
        ]);

        $options = [
            'columns' => [
                'A' => 'code',
                'B' => 'name',
                'C' => ['key' => 'val', 'type' => 'int'],
            ],
        ];

        // Warm up
        Excel::read($path, $options);
        gc_collect_cycles();
        $initialMem = memory_get_usage();

        for ($i = 0; $i < 30; $i++) {
            Excel::read($path, $options);
        }

        gc_collect_cycles();
        $finalMem = memory_get_usage();

        // Memory difference after 30 reads should be negligible (< 1 MB)
        $diff = abs($finalMem - $initialMem);
        $this->assertLessThan(1024 * 1024, $diff, 'Memory leaked across consecutive reads');
    }
    public function testAutoDetectColumnsWithoutHeaderRow(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Apple', 10],
            ['Banana', 20],
        ]);

        $res = Excel::read($path);
        $this->assertSame(2, $res->rowCount());
        $this->assertSame('Apple', $res->rows()[0]['A']);
        $this->assertSame(10, $res->rows()[0]['B']);
    }

    public function testAutoDetectColumnsWithHeaderRow(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Product', 'Qty'],
            ['Apple', 10],
        ]);

        $res = Excel::read($path, ['header_row' => 1]);
        $this->assertSame(1, $res->rowCount());
        $this->assertSame('Apple', $res->rows()[0]['Product']);
        $this->assertSame(10, $res->rows()[0]['Qty']);
    }

    public function testMaxRowsSafetyLimitExceeded(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Row 1'],
            ['Row 2'],
            ['Row 3'],
        ]);

        $this->expectException(ReaderException::class);
        $this->expectExceptionMessage('exceeds maximum allowed rows of 2');
        Excel::read($path, ['max_rows' => 2]);
    }

    public function testCsvCustomOptions(): void
    {
        $path = $this->createCsvFile("Code;Price\nSKU-1;99");

        $res = Excel::read($path, [
            'format' => 'csv',
            'header_row' => 1,
            'csv' => ['delimiter' => ';'],
            'columns' => [
                ['header' => 'Code', 'key' => 'code'],
                ['header' => 'Price', 'key' => 'price', 'type' => 'int'],
            ],
        ]);

        $this->assertSame(1, $res->rowCount());
        $this->assertSame('SKU-1', $res->rows()[0]['code']);
        $this->assertSame(99, $res->rows()[0]['price']);
    }
}
