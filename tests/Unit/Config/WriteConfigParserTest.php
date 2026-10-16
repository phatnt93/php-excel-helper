<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Config;

use ExcelHelper\Config\WriteConfigParser;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Tests\TestCase;

final class WriteConfigParserTest extends TestCase
{
    public function testParseMinimalSingleSheet(): void
    {
        $wb = WriteConfigParser::parse([
            'name' => 'Report',
            'columns' => ['code' => 'Code', 'name' => 'Name'],
            'data' => [['code' => '001', 'name' => 'Item 1']],
        ]);

        $this->assertSame('xlsx', $wb->format);
        $this->assertCount(1, $wb->sheets);
        $sheet = $wb->sheets[0];
        $this->assertSame('Report', $sheet->name);
        $this->assertCount(2, $sheet->columns);
        $this->assertSame('code', $sheet->columns[0]->key);
        $this->assertSame('Code', $sheet->columns[0]->header);
    }

    public function testParseMultiSheet(): void
    {
        $wb = WriteConfigParser::parse([
            'format' => 'xlsx',
            'sheets' => [
                ['name' => 'Sheet1', 'columns' => ['a' => 'Col A']],
                ['name' => 'Sheet2', 'columns' => ['b' => 'Col B']],
            ],
        ]);

        $this->assertCount(2, $wb->sheets);
        $this->assertSame('Sheet1', $wb->sheets[0]->name);
        $this->assertSame('Sheet2', $wb->sheets[1]->name);
    }

    public function testMixRootSheetKeysWithSheetsThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Cannot specify root sheet option "name" when "sheets" is defined');

        WriteConfigParser::parse([
            'name' => 'InvalidRoot',
            'sheets' => [
                ['name' => 'Sheet1'],
            ],
        ]);
    }

    public function testBothColumnsAndHeadersThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Cannot specify both "columns" and "headers"');

        WriteConfigParser::parse([
            'columns' => ['a' => 'A'],
            'headers' => ['A', 'B'],
        ]);
    }

    public function testFormattedTypeThrowsInWriter(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Type "formatted" is read-only and cannot be used for writing');

        WriteConfigParser::parse([
            'columns' => [
                'a' => ['header' => 'A', 'type' => 'formatted'],
            ],
        ]);
    }

    public function testCsvMultiSheetThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('CSV format only supports a single sheet');

        WriteConfigParser::parse([
            'format' => 'csv',
            'sheets' => [
                ['name' => 'S1'],
                ['name' => 'S2'],
            ],
        ]);
    }

    public function testUnknownOptionThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Unknown configuration option "invalid_opt"');

        WriteConfigParser::parse([
            'invalid_opt' => true,
        ]);
    }
}
