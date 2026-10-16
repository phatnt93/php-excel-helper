<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Integration;

use ExcelHelper\Excel;
use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\Tests\TestCase;
use Generator;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

final class WriterIntegrationTest extends TestCase
{
    public function testBasicWriteAndSaveXlsx(): void
    {
        $savePath = $this->tempPath('xlsx');

        $out = Excel::write([
            'name' => 'Employees',
            'columns' => [
                'code' => 'Mã NV',
                'name' => ['header' => 'Họ Tên', 'width' => 25],
                'salary' => ['header' => 'Lương', 'type' => 'number', 'format' => '#,##0'],
                'active' => ['header' => 'Đang làm', 'type' => 'bool'],
            ],
            'data' => [
                ['code' => 'NV001', 'name' => 'Nguyễn Văn A', 'salary' => 15000000, 'active' => true],
                ['code' => 'NV002', 'name' => 'Trần Thị B', 'salary' => 18500000, 'active' => false],
            ],
        ]);

        $writtenPath = $out->save($savePath);
        $this->assertSame($savePath, $writtenPath);
        $this->assertFileExists($savePath);

        // Inspect file with raw PhpSpreadsheet
        $spreadsheet = IOFactory::load($savePath);
        $sheet = $spreadsheet->getActiveSheet();

        $this->assertSame('Employees', $sheet->getTitle());
        $this->assertSame('Mã NV', $sheet->getCell('A1')->getValue());
        $this->assertSame('Họ Tên', $sheet->getCell('B1')->getValue());
        $this->assertSame('NV001', $sheet->getCell('A2')->getValue());
        $this->assertSame(DataType::TYPE_STRING, $sheet->getCell('A2')->getDataType());
        $this->assertSame(15000000.0, (float) $sheet->getCell('C2')->getValue());
        $this->assertSame(DataType::TYPE_NUMERIC, $sheet->getCell('C2')->getDataType());
        $this->assertSame(DataType::TYPE_BOOL, $sheet->getCell('D2')->getDataType());
        $this->assertTrue((bool) $sheet->getCell('D2')->getValue());

        $spreadsheet->disconnectWorksheets();
    }

    public function testRoundTripWriteAndRead(): void
    {
        $data = [
            ['code' => 'SKU-01', 'name' => 'Laptop', 'price' => 1250.5],
            ['code' => 'SKU-02', 'name' => 'Mouse', 'price' => 25.0],
        ];

        $formats = ['xlsx', 'ods', 'csv'];

        foreach ($formats as $fmt) {
            $savePath = $this->tempPath($fmt);

            Excel::write([
                'format' => $fmt,
                'columns' => [
                    'code' => 'Mã',
                    'name' => 'Tên',
                    'price' => ['header' => 'Giá', 'type' => 'float'],
                ],
                'data' => $data,
            ])->save($savePath);

            $readRes = Excel::read($savePath, [
                'format' => $fmt,
                'header_row' => 1,
                'columns' => [
                    ['header' => 'Mã', 'key' => 'code'],
                    ['header' => 'Tên', 'key' => 'name'],
                    ['header' => 'Giá', 'key' => 'price', 'type' => 'float'],
                ],
            ]);

            $this->assertSame(2, $readRes->rowCount(), "Failed row count for {$fmt}");
            $rows = $readRes->rows();
            $this->assertSame('SKU-01', $rows[0]['code'], "Failed code for {$fmt}");
            $this->assertSame('Laptop', $rows[0]['name'], "Failed name for {$fmt}");
            $this->assertSame(1250.5, $rows[0]['price'], "Failed price for {$fmt}");
        }
    }

    public function testMultiSheetWorkbook(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'sheets' => [
                [
                    'name' => 'Customers',
                    'columns' => ['name' => 'Tên KH'],
                    'data' => [['name' => 'Cong ty A']],
                ],
                [
                    'name' => 'Customers', // Duplicated name -> Customers (2)
                    'columns' => ['order_id' => 'Mã đơn'],
                    'data' => [['order_id' => 'DH100']],
                ],
            ],
        ])->save($savePath);

        $names = Excel::sheetNames($savePath);
        $this->assertSame(['Customers', 'Customers (2)'], $names);
    }

    public function testTitleBannerAndTotalsRow(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'title' => 'BÁO CÁO DOANH THU',
            'columns' => [
                'product' => 'Sản phẩm',
                'revenue' => ['header' => 'Doanh thu', 'type' => 'number', 'format' => '#,##0'],
            ],
            'data' => [
                ['product' => 'Gói A', 'revenue' => 1000],
                ['product' => 'Gói B', 'revenue' => 2000],
            ],
            'totals' => ['revenue' => 'sum'],
            'totals_label' => 'TỔNG CỘNG',
        ])->save($savePath);

        $spreadsheet = IOFactory::load($savePath);
        $sheet = $spreadsheet->getActiveSheet();

        // Row 1: Title merged across A1:B1
        $this->assertSame('BÁO CÁO DOANH THU', $sheet->getCell('A1')->getValue());
        $this->assertTrue($sheet->getCell('A1')->isMergeRangeValueCell());

        // Row 3: Header
        $this->assertSame('Sản phẩm', $sheet->getCell('A3')->getValue());
        $this->assertSame('Doanh thu', $sheet->getCell('B3')->getValue());

        // Rows 4-5: Data
        $this->assertSame(1000.0, (float) $sheet->getCell('B4')->getValue());
        $this->assertSame(2000.0, (float) $sheet->getCell('B5')->getValue());

        // Row 6: Totals
        $this->assertSame('TỔNG CỘNG', $sheet->getCell('A6')->getValue());
        $this->assertSame('=SUM(B4:B5)', $sheet->getCell('B6')->getValue());

        $spreadsheet->disconnectWorksheets();
    }

    public function testLayoutOptionsAndBorders(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'freeze_header' => true,
            'auto_filter' => true,
            'borders' => true,
            'zebra' => true,
            'columns' => [
                'id' => ['header' => 'ID', 'width' => 10],
                'name' => ['header' => 'Name', 'width' => 30],
            ],
            'data' => [
                ['id' => 1, 'name' => 'Alice'],
                ['id' => 2, 'name' => 'Bob'],
            ],
        ])->save($savePath);

        $spreadsheet = IOFactory::load($savePath);
        $sheet = $spreadsheet->getActiveSheet();

        $this->assertSame(10.0, (float) $sheet->getColumnDimension('A')->getWidth());
        $this->assertSame(30.0, (float) $sheet->getColumnDimension('B')->getWidth());
        $this->assertSame('A2', $sheet->getFreezePane());
        $this->assertSame('A1:B1', $sheet->getAutoFilter()->getRange());

        $spreadsheet->disconnectWorksheets();
    }

    public function testTapCallbackReceivesWorksheetAndLayout(): void
    {
        $savePath = $this->tempPath('xlsx');
        $tapped = false;

        Excel::write([
            'columns' => ['a' => 'Col A'],
            'data' => [['a' => 1], ['a' => 2]],
            'tap' => function (Worksheet $sheet, \ExcelHelper\Writer\SheetLayout $layout) use (&$tapped): void {
                $tapped = true;
                $this->assertSame(1, $layout->headerRow);
                $this->assertSame(2, $layout->dataStartRow);
                $this->assertSame(3, $layout->dataEndRow);
                $sheet->setCellValue('C1', 'Tapped');
            },
        ])->save($savePath);

        $this->assertTrue($tapped);

        $spreadsheet = IOFactory::load($savePath);
        $this->assertSame('Tapped', $spreadsheet->getActiveSheet()->getCell('C1')->getValue());
        $spreadsheet->disconnectWorksheets();
    }

    public function testGeneratorDataStreaming(): void
    {
        $savePath = $this->tempPath('xlsx');

        $generator = function (): Generator {
            for ($i = 1; $i <= 50; $i++) {
                yield ['id' => $i, 'val' => 'Row ' . $i];
            }
        };

        Excel::write([
            'columns' => ['id' => 'ID', 'val' => 'Val'],
            'data' => $generator(),
        ])->save($savePath);

        $read = Excel::read($savePath, ['columns' => ['A' => 'id', 'B' => 'val'], 'start_row' => 2]);
        $this->assertSame(50, $read->rowCount());
        $this->assertSame('Row 50', $read->rows()[49]['val']);
    }

    public function testFormulaInjectionSafety(): void
    {
        $savePath = $this->tempPath('xlsx');
        $malicious = "=cmd|' /C calc'!A0";

        Excel::write([
            'columns' => [
                'safe' => 'SafeText',
                'calc' => ['header' => 'Calc', 'type' => 'formula'],
            ],
            'data' => [
                ['safe' => $malicious, 'calc' => '1+1'],
            ],
        ])->save($savePath);

        $spreadsheet = IOFactory::load($savePath);
        $sheet = $spreadsheet->getActiveSheet();

        // Safe column is strictly TYPE_STRING, NOT TYPE_FORMULA
        $this->assertSame(DataType::TYPE_STRING, $sheet->getCell('A2')->getDataType());
        $this->assertSame($malicious, $sheet->getCell('A2')->getValue());

        // Explicit formula column is TYPE_FORMULA
        $this->assertSame(DataType::TYPE_FORMULA, $sheet->getCell('B2')->getDataType());
        $this->assertSame('=1+1', $sheet->getCell('B2')->getValue());

        $spreadsheet->disconnectWorksheets();
    }
    public function testCsvFormulaInjectionEscaping(): void
    {
        $savePath = $this->tempPath('csv');
        $malicious = "=cmd|' /C calc'!A0";

        // Default: escape_formulas is true
        Excel::write([
            'format' => 'csv',
            'columns' => ['text' => 'Text'],
            'data' => [['text' => $malicious]],
        ])->save($savePath);

        $content = (string) file_get_contents($savePath);
        $this->assertStringContainsString("\"'=cmd|", $content);

        // Explicitly disabled escape_formulas
        Excel::write([
            'format' => 'csv',
            'csv' => ['escape_formulas' => false],
            'columns' => ['text' => 'Text'],
            'data' => [['text' => $malicious]],
        ])->save($savePath);

        $rawContent = (string) file_get_contents($savePath);
        $this->assertStringContainsString('"=cmd|', $rawContent);
    }
    public function testCsvNegativeNumbersNotCorrupted(): void
    {
        $savePath = $this->tempPath('csv');

        Excel::write([
            'format' => 'csv',
            'columns' => [
                'int_neg' => 'IntNeg',
                'float_neg' => 'FloatNeg',
                'str_neg' => 'StrNeg',
            ],
            'data' => [
                ['int_neg' => -100, 'float_neg' => -5.5, 'str_neg' => '-42'],
            ],
        ])->save($savePath);

        $content = (string) file_get_contents($savePath);
        $this->assertStringNotContainsString("'-100", $content);
        $this->assertStringNotContainsString("'-5.5", $content);
        $this->assertStringNotContainsString("'-42", $content);

        $read = Excel::read($savePath, [
            'columns' => [
                'A' => ['key' => 'int_neg', 'type' => 'int'],
                'B' => ['key' => 'float_neg', 'type' => 'float'],
                'C' => ['key' => 'str_neg', 'type' => 'int'],
            ],
            'start_row' => 2,
        ]);

        $this->assertSame(-100, $read->rows()[0]['int_neg']);
        $this->assertSame(-5.5, $read->rows()[0]['float_neg']);
        $this->assertSame(-42, $read->rows()[0]['str_neg']);
    }

    public function testToStringReturnsBinaryData(): void
    {
        $binary = Excel::write([
            'columns' => ['x' => 'X'],
            'data' => [['x' => 123]],
        ])->toString();

        $this->assertNotEmpty($binary);
        // Standard XLSX zip header
        $this->assertStringStartsWith("PK\x03\x04", $binary);
    }

    public function testXlsLimitCheck(): void
    {
        // 257 columns exceeds XLS limit (256)
        $cols = [];
        for ($i = 1; $i <= 257; $i++) {
            $cols['c' . $i] = 'Col ' . $i;
        }

        $this->expectException(WriterException::class);
        $this->expectExceptionMessage('XLS (BIFF8) format limit exceeded');

        Excel::write([
            'format' => 'xls',
            'columns' => $cols,
            'data' => [],
        ])->save($this->tempPath('xls'));
    }

    public function testDirectoryNotWritableThrows(): void
    {
        $this->expectException(WriterException::class);
        $this->expectExceptionMessage('Target directory is not writable');

        Excel::write([])->save('/nonexistent_directory/path/file.xlsx');
    }
}
