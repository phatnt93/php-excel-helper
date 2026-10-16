<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Integration;

use ExcelHelper\Excel;
use ExcelHelper\Reader\ChunkContext;
use ExcelHelper\Tests\TestCase;

final class AdvancedReaderTest extends TestCase
{
    public function testChunkReadingSpreadsheet(): void
    {
        $rows = [];
        for ($i = 1; $i <= 250; $i++) {
            $rows[] = ['SKU-' . $i, $i * 10];
        }

        $path = $this->createSpreadsheetFile($rows, 'xlsx');

        $receivedChunks = [];
        $summary = Excel::chunk(
            path: $path,
            options: [
                'columns' => [
                    'A' => 'sku',
                    'B' => ['key' => 'price', 'type' => 'int'],
                ],
            ],
            size: 100,
            callback: function (array $chunkRows, ChunkContext $ctx) use (&$receivedChunks): void {
                $receivedChunks[] = [
                    'index' => $ctx->chunkIndex(),
                    'count' => count($chunkRows),
                    'start' => $ctx->startRow,
                    'end' => $ctx->endRow,
                    'processed' => $ctx->totalRowsProcessed,
                ];
            },
        );

        $this->assertSame(250, $summary->totalRows());
        $this->assertSame(3, $summary->totalChunks());
        $this->assertFalse($summary->hasErrors());

        $this->assertCount(3, $receivedChunks);
        $this->assertSame(100, $receivedChunks[0]['count']);
        $this->assertSame(100, $receivedChunks[1]['count']);
        $this->assertSame(50, $receivedChunks[2]['count']);
    }

    public function testChunkEarlyStop(): void
    {
        $rows = [];
        for ($i = 1; $i <= 200; $i++) {
            $rows[] = ['Item ' . $i];
        }

        $path = $this->createSpreadsheetFile($rows, 'xlsx');

        $processedCount = 0;
        $summary = Excel::chunk(
            path: $path,
            options: ['columns' => ['A' => 'name']],
            size: 50,
            callback: function (array $chunkRows, ChunkContext $ctx) use (&$processedCount): bool {
                $processedCount += count($chunkRows);

                // Stop after the first chunk
                return false;
            },
        );

        $this->assertSame(50, $processedCount);
        $this->assertSame(50, $summary->totalRows());
        $this->assertSame(1, $summary->totalChunks());
    }

    public function testChunkReadingCsv(): void
    {
        $csvContent = "sku,qty\n";
        for ($i = 1; $i <= 150; $i++) {
            $csvContent .= "ITEM-{$i},{$i}\n";
        }

        $path = $this->createCsvFile($csvContent);

        $total = 0;
        $summary = Excel::chunk(
            path: $path,
            options: [
                'format' => 'csv',
                'header_row' => 1,
                'columns' => [
                    ['header' => 'sku', 'key' => 'sku'],
                    ['header' => 'qty', 'key' => 'qty', 'type' => 'int'],
                ],
            ],
            size: 60,
            callback: function (array $chunkRows) use (&$total): void {
                $total += count($chunkRows);
            },
        );

        $this->assertSame(150, $total);
        $this->assertSame(150, $summary->totalRows());
        $this->assertSame(3, $summary->totalChunks());
    }

    public function testCsvAutoDelimiterDetection(): void
    {
        // Semicolon delimited CSV
        $pathSemicolon = $this->createCsvFile("Code;Name;Price\nNV1;Alice;100\nNV2;Bob;200");

        $resSemi = Excel::read($pathSemicolon, [
            'format' => 'csv',
            'header_row' => 1,
            'columns' => [
                ['header' => 'Code', 'key' => 'code'],
                ['header' => 'Name', 'key' => 'name'],
                ['header' => 'Price', 'key' => 'price', 'type' => 'int'],
            ],
        ]);

        $this->assertSame(2, $resSemi->rowCount());
        $this->assertSame('Alice', $resSemi->rows()[0]['name']);
        $this->assertSame(200, $resSemi->rows()[1]['price']);

        // Tab delimited CSV
        $pathTab = $this->createCsvFile("Code\tName\tPrice\nNV1\tAlice\t100\nNV2\tBob\t200");

        $resTab = Excel::read($pathTab, [
            'format' => 'csv',
            'header_row' => 1,
            'columns' => [
                ['header' => 'Code', 'key' => 'code'],
                ['header' => 'Name', 'key' => 'name'],
                ['header' => 'Price', 'key' => 'price', 'type' => 'int'],
            ],
        ]);

        $this->assertSame(2, $resTab->rowCount());
        $this->assertSame('Bob', $resTab->rows()[1]['name']);
    }

    public function testCsvUtf8WithBom(): void
    {
        $content = "\xEF\xBB\xBFCode,Name\nNV1,Nguyễn Văn A\n";
        $path = $this->createCsvFile($content);

        $res = Excel::read($path, [
            'format' => 'csv',
            'header_row' => 1,
            'columns' => [
                ['header' => 'Code', 'key' => 'code'],
                ['header' => 'Name', 'key' => 'name'],
            ],
        ]);

        $this->assertSame(1, $res->rowCount());
        $this->assertSame('NV1', $res->rows()[0]['code']);
        $this->assertSame('Nguyễn Văn A', $res->rows()[0]['name']);
    }

    public function testExcelSheetNamesAndInfo(): void
    {
        $path = $this->createMultiSheetFile([
            ['name' => 'FirstSheet', 'rows' => [['A', 'B'], ['1', '2']]],
            ['name' => 'SecondSheet', 'rows' => [['X', 'Y', 'Z'], ['3', '4', '5']]],
        ]);

        $names = Excel::sheetNames($path);
        $this->assertSame(['FirstSheet', 'SecondSheet'], $names);

        $info = Excel::info($path);
        $this->assertCount(2, $info);
        $this->assertSame('FirstSheet', $info[0]->name);
        $this->assertSame(2, $info[0]->totalRows);
        $this->assertSame(2, $info[0]->totalColumns);

        $this->assertSame('SecondSheet', $info[1]->name);
        $this->assertSame(2, $info[1]->totalRows);
        $this->assertSame(3, $info[1]->totalColumns);

        $infoArr = $info[0]->toArray();
        $this->assertSame('FirstSheet', $infoArr['name']);
    }

    public function testExcelHeaders(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Mã NV', 'Họ Và Tên', 'Lương'],
            ['NV01', 'An', 1000],
        ]);

        $headers = Excel::headers($path, row: 1);
        $this->assertSame('Mã NV', $headers['A']);
        $this->assertSame('Họ Và Tên', $headers['B']);
        $this->assertSame('Lương', $headers['C']);
    }
    public function testChunkInvalidSizeThrows(): void
    {
        $path = $this->createSpreadsheetFile([['Data']]);
        $this->expectException(\ExcelHelper\Exceptions\InvalidConfigException::class);
        $this->expectExceptionMessage('Chunk size must be >= 1');
        Excel::chunk($path, [], 0, static fn() => true);
    }

    public function testChunkContextStopAndErrors(): void
    {
        $path = $this->createSpreadsheetFile([
            ['10'],
            ['invalid_int'],
            ['30'],
        ]);

        $summary = Excel::chunk(
            path: $path,
            options: ['columns' => ['A' => ['key' => 'n', 'type' => 'int']]],
            size: 10,
            callback: function (array $rows, ChunkContext $ctx): void {
                $this->assertCount(1, $ctx->errors());
                $ctx->stop();
            },
        );

        $this->assertSame(1, $summary->totalChunks());
        $this->assertSame(1, $summary->errorCount());
    }

    public function testCsvPipeDelimiterAndEmptyRows(): void
    {
        $content = "A|B\n1|2\n\n3|4\n";
        $path = $this->createCsvFile($content);

        $summary = Excel::chunk(
            path: $path,
            options: [
                'format' => 'csv',
                'columns' => ['A' => 'col_a', 'B' => 'col_b'],
                'skip_empty_rows' => true,
            ],
            size: 10,
            callback: static fn() => true,
        );

        $this->assertSame(3, $summary->totalRows());
    }

    public function testCsvChunkMaxRowsExceeded(): void
    {
        $content = "A\n1\n2\n3\n4\n5\n";
        $path = $this->createCsvFile($content);

        $this->expectException(\ExcelHelper\Exceptions\ReaderException::class);
        $this->expectExceptionMessage('exceeds maximum allowed rows of 2');

        Excel::chunk(
            path: $path,
            options: ['format' => 'csv', 'max_rows' => 2],
            size: 2,
            callback: static fn() => true,
        );
    }

    public function testCsvChunkOnErrorThrowAndSkip(): void
    {
        $content = "val\n10\ninvalid\n30\n";
        $path = $this->createCsvFile($content);

        // Skip row
        $skippedRows = [];
        $summary = Excel::chunk(
            path: $path,
            options: [
                'format' => 'csv',
                'header_row' => 1,
                'on_error' => 'skip_row',
                'columns' => [['header' => 'val', 'key' => 'n', 'type' => 'int']],
            ],
            size: 10,
            callback: function (array $rows) use (&$skippedRows): void {
                $skippedRows = array_merge($skippedRows, $rows);
            },
        );

        $this->assertCount(2, $skippedRows);
        $this->assertSame(10, $skippedRows[0]['n']);
        $this->assertSame(30, $skippedRows[1]['n']);

        // Throw
        $this->expectException(\ExcelHelper\Exceptions\RowValidationException::class);
        Excel::chunk(
            path: $path,
            options: [
                'format' => 'csv',
                'header_row' => 1,
                'on_error' => 'throw',
                'columns' => [['header' => 'val', 'key' => 'n', 'type' => 'int']],
            ],
            size: 10,
            callback: static fn() => true,
        );
    }

    public function testCsvUtf16BomDetection(): void
    {
        // UTF-16LE with BOM
        $content = "\xFF\xFE" . mb_convert_encoding("Code,Name\nA1,Item1\n", 'UTF-16LE', 'UTF-8');
        $path = $this->createCsvFile($content);

        $summary = Excel::chunk(
            path: $path,
            options: [
                'format' => 'csv',
                'header_row' => 1,
                'columns' => [
                    ['header' => 'Code', 'key' => 'c'],
                    ['header' => 'Name', 'key' => 'n'],
                ],
            ],
            size: 10,
            callback: static fn() => true,
        );

        $this->assertSame(1, $summary->totalRows());
    }
}
