<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit;

use ExcelHelper\Excel;
use ExcelHelper\Tests\TestCase;

final class ExcelTest extends TestCase
{
    public function testSetAndGetLocale(): void
    {
        Excel::setLocale('vi');
        $this->assertSame('vi', Excel::getLocale());

        Excel::resetDefaults();
        $this->assertSame('en', Excel::getLocale());
    }

    public function testReadNonExistentThrowsReaderException(): void
    {
        $this->expectException(\ExcelHelper\Exceptions\ReaderException::class);
        Excel::read('dummy.xlsx');
    }

    public function testReaderReturnsReaderBuilder(): void
    {
        $reader = Excel::reader('dummy.xlsx');
        $this->assertInstanceOf(\ExcelHelper\Reader\ReaderBuilder::class, $reader);
    }

    public function testWriteReturnsExcelOutput(): void
    {
        $out = Excel::write([]);
        $this->assertInstanceOf(\ExcelHelper\Writer\ExcelOutput::class, $out);
    }

    public function testWriterReturnsWriterBuilder(): void
    {
        $writer = Excel::writer();
        $this->assertInstanceOf(\ExcelHelper\Writer\WriterBuilder::class, $writer);
    }
}
