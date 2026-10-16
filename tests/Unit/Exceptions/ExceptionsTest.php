<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Exceptions;

use ExcelHelper\Config\ConfigPath;
use ExcelHelper\Exceptions\ExcelHelperException;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\RowValidationException;
use ExcelHelper\Exceptions\SheetNotFoundException;
use ExcelHelper\Exceptions\UnsupportedFormatException;
use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\Reader\RowError;
use ExcelHelper\Tests\TestCase;
use RuntimeException;

final class ExceptionsTest extends TestCase
{
    public function testInvalidConfigException(): void
    {
        $path = ConfigPath::root('test')->append('option');
        $e = InvalidConfigException::atPath($path, 'Value must be int.');

        $this->assertInstanceOf(ExcelHelperException::class, $e);
        $this->assertSame('Configuration error at "test.option": Value must be int.', $e->getMessage());
        $this->assertSame($path, $e->getConfigPath());
    }

    public function testRowValidationException(): void
    {
        $error = new RowError(10, 'C', 'amount', 'invalid_int', 'foo', 'Must be int.');
        $e = RowValidationException::fromRowError($error);

        $this->assertInstanceOf(ExcelHelperException::class, $e);
        $this->assertSame($error, $e->getRowError());
        $this->assertStringContainsString('Row 10 column "C" (amount) validation failed [invalid_int]: Must be int.', $e->getMessage());
        $this->assertSame([
            'row' => 10,
            'column' => 'C',
            'key' => 'amount',
            'code' => 'invalid_int',
            'value' => 'foo',
            'message' => 'Must be int.',
        ], $error->toArray());
    }

    public function testReaderExceptions(): void
    {
        $e1 = ReaderException::fileNotFound('/path/missing.xlsx');
        $this->assertStringContainsString('/path/missing.xlsx', $e1->getMessage());

        $e2 = ReaderException::notReadable('/path/unreadable.xlsx');
        $this->assertStringContainsString('not readable', $e2->getMessage());

        $e3 = ReaderException::isADirectory('/path/dir');
        $this->assertStringContainsString('is a directory', $e3->getMessage());

        $e4 = ReaderException::maxFileSizeExceeded('/path/file.xlsx', 1000, 2000);
        $this->assertStringContainsString('exceeds maximum allowed size', $e4->getMessage());

        $e5 = ReaderException::maxRowsExceeded(100, 200);
        $this->assertStringContainsString('exceeds maximum allowed rows', $e5->getMessage());

        $e6 = ReaderException::missingHeaderColumn('Email');
        $this->assertStringContainsString('Required header column "Email"', $e6->getMessage());

        $e7 = ReaderException::cannotOpenFile('/path/file.xlsx');
        $this->assertStringContainsString('Cannot open spreadsheet', $e7->getMessage());

        $e8 = ReaderException::formulaError('A1', 'Division by zero');
        $this->assertStringContainsString('cell "A1"', $e8->getMessage());

        $e9 = ReaderException::emptyFile('/path/empty.xlsx');
        $this->assertStringContainsString('is empty', $e9->getMessage());

        $e10 = ReaderException::corruptedFile('/path/corrupt.xlsx');
        $this->assertStringContainsString('is corrupted', $e10->getMessage());

        $e11 = ReaderException::withMessage('Custom reader error');
        $this->assertStringContainsString('Custom reader error', $e11->getMessage());
    }

    public function testUnsupportedFormatException(): void
    {
        $e1 = UnsupportedFormatException::notWhitelisted('exe', ['xlsx', 'csv']);
        $this->assertStringContainsString('not in the allowed formats whitelist', $e1->getMessage());

        $e2 = UnsupportedFormatException::unrecognized('/path/unknown.bin');
        $this->assertStringContainsString('Could not recognize', $e2->getMessage());

        $e3 = UnsupportedFormatException::unsupportedForReading('pdf');
        $this->assertStringContainsString('not supported for reading', $e3->getMessage());

        $e4 = UnsupportedFormatException::unsupportedForWriting('pdf');
        $this->assertStringContainsString('not supported for writing', $e4->getMessage());
    }

    public function testSheetNotFoundException(): void
    {
        $e1 = SheetNotFoundException::byIndex(2, ['Sheet1', 'Sheet2']);
        $this->assertStringContainsString('index 2', $e1->getMessage());

        $e2 = SheetNotFoundException::byName('Users', ['Sheet1', 'Sheet2']);
        $this->assertStringContainsString('name "Users"', $e2->getMessage());
    }

    public function testWriterExceptions(): void
    {
        $e1 = WriterException::directoryNotWritable('/protected');
        $this->assertStringContainsString('not writable', $e1->getMessage());

        $e2 = WriterException::headersAlreadySent();
        $this->assertStringContainsString('headers have already been sent', $e2->getMessage());

        $e3 = WriterException::unsupportedTypeForWriting('formatted', 'read only');
        $this->assertStringContainsString('cannot be used for writing', $e3->getMessage());

        $e4 = WriterException::xlsLimitExceeded(70000, 300);
        $this->assertStringContainsString('limit exceeded', $e4->getMessage());

        $prev = new RuntimeException('Disk full');
        $e5 = WriterException::failedToSave('/path/out.xlsx', $prev);
        $this->assertSame($prev, $e5->getPrevious());

        $e6 = WriterException::cannotWriteValue('B2', 'Cannot convert object');
        $this->assertStringContainsString('cell "B2"', $e6->getMessage());

        $e7 = WriterException::withMessage('Generic writer error');
        $this->assertSame('Generic writer error', $e7->getMessage());
    }
}
