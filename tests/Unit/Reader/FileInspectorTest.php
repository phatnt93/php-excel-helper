<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Reader;

use ExcelHelper\Config\ReadConfigParser;
use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\UnsupportedFormatException;
use ExcelHelper\Reader\FileInspector;
use ExcelHelper\Tests\TestCase;

final class FileInspectorTest extends TestCase
{
    public function testFileNotFoundThrows(): void
    {
        $opts = ReadConfigParser::parse([]);
        $this->expectException(ReaderException::class);
        $this->expectExceptionMessage('Spreadsheet file not found');
        FileInspector::inspect('/nonexistent/path/file.xlsx', $opts);
    }

    public function testDirectoryThrows(): void
    {
        $opts = ReadConfigParser::parse([]);
        $this->expectException(ReaderException::class);
        $this->expectExceptionMessage('Path is a directory, not a file');
        FileInspector::inspect(sys_get_temp_dir(), $opts);
    }

    public function testEmptyFileThrows(): void
    {
        $emptyPath = $this->tempPath('xlsx');
        touch($emptyPath);

        $opts = ReadConfigParser::parse([]);
        $this->expectException(ReaderException::class);
        $this->expectExceptionMessage('Spreadsheet file is empty');
        FileInspector::inspect($emptyPath, $opts);
    }

    public function testMaxFileSizeExceededThrows(): void
    {
        $path = $this->createSpreadsheetFile([['Header'], ['Value']]);

        $opts = ReadConfigParser::parse(['max_file_size' => 10]); // 10 bytes limit
        $this->expectException(ReaderException::class);
        $this->expectExceptionMessage('exceeds maximum allowed size');
        FileInspector::inspect($path, $opts);
    }

    public function testFormatNotWhitelistedThrows(): void
    {
        $path = $this->createSpreadsheetFile([['Header'], ['Value']], 'xlsx');

        $opts = ReadConfigParser::parse(['allowed_formats' => ['csv']]);
        $this->expectException(UnsupportedFormatException::class);
        $this->expectExceptionMessage('not in the allowed formats whitelist');
        FileInspector::inspect($path, $opts);
    }

    public function testFakeSpreadsheetRejected(): void
    {
        $fakePath = $this->createFakeSpreadsheet('xlsx');

        $opts = ReadConfigParser::parse([]);
        $this->expectException(UnsupportedFormatException::class);
        $this->expectExceptionMessage('not in the allowed formats whitelist');
        FileInspector::inspect($fakePath, $opts);
    }

    public function testValidXlsxInspection(): void
    {
        $path = $this->createSpreadsheetFile([['Header'], ['Value']], 'xlsx');

        $opts = ReadConfigParser::parse([]);
        $res = FileInspector::inspect($path, $opts);

        $this->assertSame('Xlsx', $res['readerType']);
        $this->assertSame('xlsx', $res['format']);
    }
}
