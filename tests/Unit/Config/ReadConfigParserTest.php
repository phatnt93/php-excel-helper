<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Config;

use ExcelHelper\Config\ReadConfigParser;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Tests\TestCase;

final class ReadConfigParserTest extends TestCase
{
    public function testParseMinimalDefaultConfig(): void
    {
        $opts = ReadConfigParser::parse([]);

        $this->assertSame('auto', $opts->format);
        $this->assertSame(0, $opts->sheet);
        $this->assertNull($opts->headerRow);
        $this->assertSame(1, $opts->startRow);
        $this->assertNull($opts->endRow);
        $this->assertSame('A', $opts->startColumn);
        $this->assertSame([], $opts->columns);
        $this->assertTrue($opts->skipEmptyRows);
        $this->assertFalse($opts->stopAtEmptyRow);
        $this->assertTrue($opts->trim);
        $this->assertSame('collect', $opts->onError);
        $this->assertNull($opts->rowNumberKey);
        $this->assertFalse($opts->calculateFormulas);
        $this->assertNull($opts->maxRows);
        $this->assertNull($opts->maxFileSize);
    }

    public function testStartRowDefaultsToHeaderRowPlusOne(): void
    {
        $opts = ReadConfigParser::parse(['header_row' => 3]);
        $this->assertSame(3, $opts->headerRow);
        $this->assertSame(4, $opts->startRow);
    }

    public function testParseKeyedColumns(): void
    {
        $opts = ReadConfigParser::parse([
            'columns' => [
                'A' => 'code',
                'B' => [
                    'key' => 'name',
                    'type' => 'string',
                    'required' => true,
                ],
            ],
        ]);

        $this->assertCount(2, $opts->columns);
        $this->assertSame('code', $opts->columns[0]->key);
        $this->assertSame(1, $opts->columns[0]->column);

        $this->assertSame('name', $opts->columns[1]->key);
        $this->assertSame(2, $opts->columns[1]->column);
        $this->assertTrue($opts->columns[1]->required);
    }

    public function testParseSequentialColumns(): void
    {
        $opts = ReadConfigParser::parse([
            'start_column' => 'B',
            'columns' => [
                'code',
                null, // Skip column C
                ['key' => 'salary', 'type' => 'float'],
            ],
        ]);

        $this->assertCount(2, $opts->columns);
        $this->assertSame('code', $opts->columns[0]->key);
        $this->assertSame(2, $opts->columns[0]->column); // B = 2

        $this->assertSame('salary', $opts->columns[1]->key);
        $this->assertSame(4, $opts->columns[1]->column); // D = 4 (C skipped)
    }

    public function testParseHeaderColumns(): void
    {
        $opts = ReadConfigParser::parse([
            'header_row' => 1,
            'columns' => [
                ['header' => 'Mã NV', 'key' => 'code'],
                ['header' => ['Email', 'E-mail'], 'key' => 'email', 'optional' => true],
            ],
        ]);

        $this->assertCount(2, $opts->columns);
        $this->assertSame('code', $opts->columns[0]->key);
        $this->assertSame('Mã NV', $opts->columns[0]->header);

        $this->assertSame('email', $opts->columns[1]->key);
        $this->assertSame(['Email', 'E-mail'], $opts->columns[1]->header);
        $this->assertTrue($opts->columns[1]->optional);
    }

    public function testUnknownOptionThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Unknown configuration option "strat_row"');
        ReadConfigParser::parse(['strat_row' => 2]);
    }

    public function testStartRowLessThanOneThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Start row must be >= 1');
        ReadConfigParser::parse(['start_row' => 0]);
    }

    public function testStartRowLessThanHeaderRowThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Start row (2) must be greater than header row (3)');
        ReadConfigParser::parse(['header_row' => 3, 'start_row' => 2]);
    }

    public function testEndRowLessThanStartRowThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('End row (5) cannot be less than start row (10)');
        ReadConfigParser::parse(['start_row' => 10, 'end_row' => 5]);
    }

    public function testMixSequentialAndHeaderColumnsThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Cannot mix sequential column definitions with header-based');
        ReadConfigParser::parse([
            'header_row' => 1,
            'columns' => [
                'code',
                ['header' => 'Email', 'key' => 'email'],
            ],
        ]);
    }

    public function testDuplicateColumnKeysThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Duplicate column key "code"');
        ReadConfigParser::parse([
            'columns' => [
                'A' => 'code',
                'B' => 'code',
            ],
        ]);
    }

    public function testDuplicateColumnPositionsThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Duplicate column');
        ReadConfigParser::parse([
            'columns' => [
                'A' => 'code1',
                1 => 'code2',
            ],
        ]);
    }

    public function testInvalidOnErrorThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Invalid on_error "explode"');
        ReadConfigParser::parse(['on_error' => 'explode']);
    }

    public function testUnknownColumnTypeThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Unknown column type "imaginary_type"');
        ReadConfigParser::parse([
            'columns' => [
                'A' => ['key' => 'x', 'type' => 'imaginary_type'],
            ],
        ]);
    }

    public function testHeaderColumnsWithoutHeaderRowThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Header-based column definitions require "header_row"');
        ReadConfigParser::parse([
            'columns' => [
                ['header' => 'Name', 'key' => 'name'],
            ],
        ]);
    }
}
