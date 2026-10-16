<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Integration;

use ExcelHelper\Config\ReadConfigParser;
use ExcelHelper\Excel;
use ExcelHelper\Reader\FastXlsxStreamReader;
use ExcelHelper\Tests\TestCase;

final class FastXlsxStreamReaderTest extends TestCase
{
    public function testStreamXlsxGeneratorMemoryO1(): void
    {
        $savePath = $this->tempPath('xlsx');

        // Create a 500-row XLSX spreadsheet
        $data = [];
        for ($i = 1; $i <= 500; $i++) {
            $data[] = [
                'id' => $i,
                'name' => 'User ' . $i,
                'score' => $i * 1.5,
            ];
        }

        Excel::write([
            'columns' => ['id' => 'ID', 'name' => 'Name', 'score' => 'Score'],
            'data' => $data,
        ])->save($savePath);

        $initialMem = memory_get_usage();
        $count = 0;
        $firstRow = null;
        $lastRow = null;

        foreach (Excel::stream($savePath, ['columns' => ['A' => ['key' => 'id', 'type' => 'int'], 'B' => 'name', 'C' => ['key' => 'score', 'type' => 'float']], 'start_row' => 2]) as $row) {
            $count++;
            if ($firstRow === null) {
                $firstRow = $row;
            }
            $lastRow = $row;
        }

        $memDiff = memory_get_usage() - $initialMem;

        $this->assertSame(500, $count);
        $this->assertNotNull($firstRow);
        $this->assertSame(1, $firstRow['id']);
        $this->assertSame('User 1', $firstRow['name']);
        $this->assertSame(1.5, $firstRow['score']);

        $this->assertNotNull($lastRow);
        $this->assertSame(500, $lastRow['id']);
        $this->assertSame('User 500', $lastRow['name']);
        $this->assertSame(750.0, $lastRow['score']);

        // Memory difference during streaming should be minimal (< 2MB)
        $this->assertLessThan(2 * 1024 * 1024, $memDiff);
    }

    public function testFastXlsxStreamReaderReadResult(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'headers' => ['Mã', 'Tên sản phẩm', 'Giá'],
            'data' => [
                ['SP01', 'Bàn phím cơ', 1500000],
                ['SP02', 'Chuột không dây', 450000],
            ],
        ])->save($savePath);

        $opts = ReadConfigParser::parse([
            'header_row' => 1,
            'columns' => [
                ['header' => 'Mã', 'key' => 'code'],
                ['header' => 'Tên sản phẩm', 'key' => 'name'],
                ['header' => 'Giá', 'key' => 'price', 'type' => 'number'],
            ],
        ]);

        $result = FastXlsxStreamReader::read($savePath, $opts);
        $this->assertSame(2, $result->rowCount());
        $this->assertSame('SP01', $result->rows()[0]['code']);
        $this->assertSame('Bàn phím cơ', $result->rows()[0]['name']);
        $this->assertEquals(1500000, $result->rows()[0]['price']);
    }
    public function testFastXlsxEmptyRowsAndStopAtEmpty(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'columns' => ['val' => 'Val'],
            'data' => [
                ['val' => 'Row 1'],
                ['val' => null], // empty row
                ['val' => 'Row 3'],
            ],
        ])->save($savePath);

        // skip_empty_rows default true
        $res = FastXlsxStreamReader::read($savePath, ReadConfigParser::parse([
            'columns' => ['A' => 'val'],
            'start_row' => 2,
        ]));
        $this->assertSame(2, $res->rowCount());

        // stop_at_empty_row
        $resStop = FastXlsxStreamReader::read($savePath, ReadConfigParser::parse([
            'columns' => ['A' => 'val'],
            'start_row' => 2,
            'stop_at_empty_row' => true,
        ]));
        $this->assertSame(1, $resStop->rowCount());
        $this->assertSame('Row 1', $resStop->rows()[0]['val']);
    }

    public function testFastXlsxOnErrorModes(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'columns' => ['age' => 'Age'],
            'data' => [
                ['age' => 'not_an_int'],
                ['age' => 25],
            ],
        ])->save($savePath);

        // on_error => collect (default)
        $resCollect = FastXlsxStreamReader::read($savePath, ReadConfigParser::parse([
            'columns' => ['A' => ['key' => 'age', 'type' => 'int']],
            'start_row' => 2,
        ]));
        $this->assertSame(2, $resCollect->rowCount());
        $this->assertCount(1, $resCollect->errors());

        // on_error => skip_row
        $resSkip = FastXlsxStreamReader::read($savePath, ReadConfigParser::parse([
            'columns' => ['A' => ['key' => 'age', 'type' => 'int']],
            'start_row' => 2,
            'on_error' => 'skip_row',
        ]));
        $this->assertSame(1, $resSkip->rowCount());
        $this->assertSame(25, $resSkip->rows()[0]['age']);

        // on_error => throw
        $this->expectException(\ExcelHelper\Exceptions\RowValidationException::class);
        FastXlsxStreamReader::read($savePath, ReadConfigParser::parse([
            'columns' => ['A' => ['key' => 'age', 'type' => 'int']],
            'start_row' => 2,
            'on_error' => 'throw',
        ]));
    }

    public function testFastXlsxAutoDetectedColumns(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'show_header' => false,
            'data' => [
                ['A', 'B', 'C'],
                [10, 20, 30],
            ],
        ])->save($savePath);

        $res = FastXlsxStreamReader::read($savePath, ReadConfigParser::parse([
            'header_row' => 1,
            'start_row' => 2,
        ]));
        $this->assertSame(1, $res->rowCount());
        $this->assertSame(10, $res->rows()[0]['A']);
    }
}
