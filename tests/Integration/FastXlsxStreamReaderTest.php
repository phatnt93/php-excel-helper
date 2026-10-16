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
    public function testFastXlsxInlineStrings(): void
    {
        $savePath = $this->tempPath('xlsx');

        $ss = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $ss->getActiveSheet();
        $sheet->setCellValue('A1', 'Placeholder');
        $sheet->setCellValue('B1', 'Normal String');
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($ss);
        $writer->save($savePath);
        $ss->disconnectWorksheets();

        $zip = new \ZipArchive();
        $zip->open($savePath);
        $xml = (string) $zip->getFromName('xl/worksheets/sheet1.xml');
        $xml = preg_replace(
            '/<c r="A1"[^>]*>.*?<\/c>/s',
            '<c r="A1" t="inlineStr"><is><t>Inline Text Value</t></is></c>',
            $xml,
        );
        $zip->addFromString('xl/worksheets/sheet1.xml', (string) $xml);
        $zip->close();

        $result = FastXlsxStreamReader::read($savePath, ReadConfigParser::parse([
            'columns' => [
                'A' => 'inline',
                'B' => 'shared',
            ],
            'start_row' => 1,
        ]));

        $this->assertSame(1, $result->rowCount());
        $this->assertSame('Inline Text Value', $result->rows()[0]['inline']);
        $this->assertSame('Normal String', $result->rows()[0]['shared']);
    }

    public function testCsvStreamLazinessAndBoundedMemory(): void
    {
        $csvPath = $this->tempPath('csv');
        $fp = fopen($csvPath, 'w');
        if ($fp !== false) {
            fputcsv($fp, ['id', 'name', 'value']);
            for ($i = 1; $i <= 5000; $i++) {
                fputcsv($fp, [$i, 'User ' . $i, 'Val ' . $i]);
            }
            fclose($fp);
        }

        $csvSmallPath = $this->tempPath('csv');
        $fpSmall = fopen($csvSmallPath, 'w');
        if ($fpSmall !== false) {
            fputcsv($fpSmall, ['id', 'name', 'value']);
            for ($i = 1; $i <= 10; $i++) {
                fputcsv($fpSmall, [$i, 'User ' . $i, 'Val ' . $i]);
            }
            fclose($fpSmall);
        }

        // Test 1: Laziness - verify only 3 rows are evaluated using a transform counter
        $transformCount = 0;
        $generator = Excel::stream($csvPath, [
            'columns' => [
                'A' => 'id',
                'B' => [
                    'key' => 'name',
                    'transform' => function ($val) use (&$transformCount) {
                        $transformCount++;
                        return $val;
                    },
                ],
                'C' => 'value',
            ],
            'start_row' => 2,
        ]);

        $this->assertInstanceOf(\Generator::class, $generator);
        $this->assertTrue($generator->valid());

        $readRows = [];
        foreach ($generator as $lineNum => $row) {
            $readRows[] = $row;
            if (count($readRows) === 3) {
                break;
            }
        }

        // True laziness: transform was invoked EXACTLY 3 times, NOT 5000 times!
        $this->assertSame(3, $transformCount);
        $this->assertCount(3, $readRows);
        $this->assertSame('User 1', $readRows[0]['name']);
        $this->assertSame('User 3', $readRows[2]['name']);

        // Test 2: Bounded memory: baseline captured BEFORE generator creation
        // Warmup call to load class definitions and avoid autoloader skew
        $warmup = Excel::stream($csvSmallPath, ['start_row' => 2]);
        $warmup->rewind();
        unset($warmup);
        gc_collect_cycles();

        // Measure small CSV (10 rows) at first yield
        $baseSmall = memory_get_usage();
        $genSmall = Excel::stream($csvSmallPath, ['start_row' => 2]);
        $genSmall->rewind();
        $memSmall = memory_get_usage() - $baseSmall;
        unset($genSmall);
        gc_collect_cycles();

        // Measure large CSV (5000 rows) at first yield
        $baseLarge = memory_get_usage();
        $genLarge = Excel::stream($csvPath, ['start_row' => 2]);
        $genLarge->rewind();
        $memLarge = memory_get_usage() - $baseLarge;

        // Allocation at first yield is strictly bounded (< 100 KB, actual ~18 KB)
        $this->assertLessThan(100 * 1024, $memLarge);
        // Allocation difference between 10 rows and 5000 rows is negligible (< 20 KB, actual 0 bytes)
        $this->assertLessThan(20 * 1024, abs($memLarge - $memSmall));

        $totalIterated = 0;
        foreach ($genLarge as $row) {
            $totalIterated++;
        }
        $this->assertSame(5000, $totalIterated);
        // Test 3: Streaming with header_row, end_row, and skip_row
        $csvErrorPath = $this->tempPath('csv');
        file_put_contents($csvErrorPath, "ID,Age\n1,invalid_int\n2,30\n\n3,40\n");
        $streamed = iterator_to_array(Excel::stream($csvErrorPath, [
            'header_row' => 1,
            'start_row' => 2,
            'end_row' => 4,
            'columns' => [
                ['header' => 'ID', 'key' => 'id'],
                ['header' => 'Age', 'key' => 'age', 'type' => 'int'],
            ],
            'skip_empty_rows' => true,
            'on_error' => 'skip_row',
        ]));
        $this->assertCount(2, $streamed);
        $this->assertSame(30, array_values($streamed)[0]['age']);
        $this->assertSame(40, array_values($streamed)[1]['age']);
    }
}
