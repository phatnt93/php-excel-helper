<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Reader;

use ExcelHelper\Excel;
use ExcelHelper\Reader\ChunkContext;
use ExcelHelper\Reader\ReaderBuilder;
use ExcelHelper\Tests\TestCase;

final class ReaderBuilderTest extends TestCase
{
    public function testFluentBuilderOptionsParity(): void
    {
        $builder = new ReaderBuilder('/data/file.xlsx');
        $builder
            ->sheet('Sheet1')
            ->headerRow(1)
            ->startRow(2)
            ->endRow(10)
            ->startColumn('B')
            ->column('B', 'code')
            ->column('C', 'name', required: true)
            ->skipEmptyRows(true)
            ->stopAtEmptyRow(false)
            ->trim(true)
            ->onError('skip_row')
            ->rowNumberKey('_row')
            ->calculateFormulas(true)
            ->maxRows(500)
            ->maxFileSize('10M')
            ->dateFormat('Y/m/d')
            ->datetimeFormat('Y/m/d H:i')
            ->csv(['delimiter' => ','])
            ->locale('vi')
            ->format('xlsx')
            ->allowedFormats(['xlsx', 'csv']);

        $opts = $builder->toOptions();

        $this->assertSame('Sheet1', $opts['sheet']);
        $this->assertSame(1, $opts['header_row']);
        $this->assertSame(2, $opts['start_row']);
        $this->assertSame(10, $opts['end_row']);
        $this->assertSame('B', $opts['start_column']);
        $this->assertArrayHasKey('B', $opts['columns']);
        $this->assertArrayHasKey('C', $opts['columns']);
        $this->assertSame('skip_row', $opts['on_error']);
        $this->assertSame('_row', $opts['row_number_key']);
        $this->assertTrue($opts['calculate_formulas']);
        $this->assertSame(500, $opts['max_rows']);
        $this->assertSame('10M', $opts['max_file_size']);
        $this->assertSame('vi', $opts['locale']);
    }

    public function testFluentReadExecution(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Mã', 'Tên', 'Lương'],
            ['NV01', 'An', 1000],
            ['NV02', 'Bình', 2000],
        ]);

        $res = Excel::reader($path)
            ->headerRow(1)
            ->startRow(2)
            ->column('A', 'code')
            ->column('B', 'name')
            ->column('C', 'salary', type: 'int')
            ->read();

        $this->assertSame(2, $res->rowCount());
        $this->assertSame('NV01', $res->rows()[0]['code']);
        $this->assertSame('An', $res->rows()[0]['name']);
        $this->assertSame(1000, $res->rows()[0]['salary']);
    }

    public function testFluentChunkExecution(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Item 1'],
            ['Item 2'],
            ['Item 3'],
        ]);

        $chunks = 0;
        $summary = Excel::reader($path)
            ->columns(['A' => 'name'])
            ->chunk(2, function (array $rows, ChunkContext $ctx) use (&$chunks): void {
                $chunks++;
            });

        $this->assertSame(3, $summary->totalRows());
        $this->assertSame(2, $chunks);
    }
}
