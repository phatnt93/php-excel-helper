<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Writer;

use ExcelHelper\Excel;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Writer\SheetBuilder;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

final class WriterBuilderTest extends TestCase
{
    public function testFluentMultiSheetWriter(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::writer()
            ->properties(['creator' => 'HR System', 'title' => 'Payroll'])
            ->format('xlsx')
            ->default('zebra', true)
            ->sheet('Nhân viên', function (SheetBuilder $s): void {
                $s->title('DANH SÁCH NHÂN VIÊN')
                    ->column('code', 'Mã NV')
                    ->column('name', 'Họ tên', width: 28)
                    ->column('salary', 'Lương', type: 'number', format: '#,##0')
                    ->data([
                        ['code' => 'NV1', 'name' => 'Alice', 'salary' => 1500],
                        ['code' => 'NV2', 'name' => 'Bob', 'salary' => 2500],
                    ])
                    ->totals(['salary' => 'sum']);
            })
            ->sheet('Phòng ban', function (SheetBuilder $s): void {
                $s->columns(['id' => 'ID', 'name' => 'Tên'])
                    ->data([['id' => 'D1', 'name' => 'Kỹ thuật']]);
            })
            ->save($savePath);

        $this->assertFileExists($savePath);
        $names = Excel::sheetNames($savePath);
        $this->assertSame(['Nhân viên', 'Phòng ban'], $names);
    }

    public function testFluentSingleSheetWriter(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::writer()
            ->title('Báo Cáo')
            ->startCell('A1')
            ->column('product', 'Sản Phẩm')
            ->column('qty', 'Số Lượng', type: 'int')
            ->data([['product' => 'Bút', 'qty' => 10]])
            ->freezeHeader(true)
            ->autoFilter(true)
            ->autoSize(true, 1000)
            ->save($savePath);

        $this->assertFileExists($savePath);
        $res = Excel::read($savePath, [
            'header_row' => 3,
            'columns' => [
                ['header' => 'Sản Phẩm', 'key' => 'product'],
                ['header' => 'Số Lượng', 'key' => 'qty', 'type' => 'int'],
            ],
        ]);

        $this->assertSame(1, $res->rowCount());
        $this->assertSame('Bút', $res->rows()[0]['product']);
        $this->assertSame(10, $res->rows()[0]['qty']);
    }

    public function testWriterBuilderOutputs(): void
    {
        $builder = Excel::writer()
            ->columns(['x' => 'X'])
            ->data([['x' => 1]]);

        $this->assertInstanceOf(Spreadsheet::class, $builder->spreadsheet());
        $this->assertNotEmpty($builder->toString());
    }
    public function testAllFluentMethodsOnWriterBuilderAndSheetBuilder(): void
    {
        $builder = Excel::writer();
        $builder
            ->format('xlsx')
            ->property('creator', 'Admin')
            ->default('auto_size', true)
            ->defaults(['zebra' => true])
            ->csv(['delimiter' => ';'])
            ->title('Title', ['bold' => true])
            ->startCell('A1')
            ->headers(['Col1', 'Col2'])
            ->headerStyle(['bold' => true])
            ->rowStyle(['italic' => true])
            ->zebra(['color' => '#EEEEEE'])
            ->borders('thin')
            ->freezeHeader(true)
            ->autoFilter(true)
            ->autoSize(true, 500)
            ->rowHeight(25.0)
            ->merge('C1:D1')
            ->merges(['E1:F1'])
            ->tap(static fn() => true)
            ->totals(['Col1' => 'sum'], 'Total')
            ->data([['Col1' => 10, 'Col2' => 20]]);

        $config = $builder->toConfig();
        $this->assertSame('xlsx', $config['format']);
        $this->assertSame('Admin', $config['properties']['creator']);
        $this->assertSame('A1', $config['start_cell']);
        $this->assertSame(25.0, $config['row_height']);
        $this->assertCount(2, $config['merges']);

        // Test SheetBuilder methods directly
        $sheetBuilder = new SheetBuilder('TestSheet');
        $sheetBuilder
            ->name('CustomName')
            ->title('My Title', ['bold' => true])
            ->startCell('B2')
            ->headers(['H1'])
            ->headerStyle(['bold' => true])
            ->rowStyle(['italic' => true])
            ->zebra(true)
            ->borders('double')
            ->freezeHeader(false)
            ->autoFilter(false)
            ->autoSize(false, 100)
            ->rowHeight(20.0)
            ->merge('A1:B1')
            ->merges(['C1:D1'])
            ->tap(static fn() => true)
            ->totals(['H1' => 'count'], 'Count')
            ->data([['H1' => 'val']]);

        $sConfig = $sheetBuilder->toConfig();
        $this->assertSame('CustomName', $sConfig['name']);
        $this->assertSame('B2', $sConfig['start_cell']);
        $this->assertSame(20.0, $sConfig['row_height']);
    }
}
