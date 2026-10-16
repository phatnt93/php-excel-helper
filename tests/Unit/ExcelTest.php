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
    public function testTemplateReturnsTemplateProcessor(): void
    {
        $path = $this->tempPath('xlsx');
        $ss = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($ss, 'Xlsx')->save($path);
        $ss->disconnectWorksheets();

        $tpl = Excel::template($path);
        $this->assertInstanceOf(\ExcelHelper\Template\TemplateProcessor::class, $tpl);
    }

    public function testStreamReturnsGenerator(): void
    {
        $path = $this->tempPath('csv');
        file_put_contents($path, "A,B\n1,2\n");

        $gen = Excel::stream($path);
        $this->assertInstanceOf(\Generator::class, $gen);
    }
}
