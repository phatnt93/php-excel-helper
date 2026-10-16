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

    public function testStreamAdvancesGeneratorAndExplicitFormatNoExtension(): void
    {
        $path = $this->tempPath('csv');
        file_put_contents($path, "A,B\n10,20\n30,40\n");

        $gen = Excel::stream($path, ['start_row' => 2]);
        $this->assertInstanceOf(\Generator::class, $gen);
        $rows = iterator_to_array($gen);
        $this->assertCount(2, $rows);
        $this->assertSame('10', $rows[2]['A']);
        $this->assertSame('20', $rows[2]['B']);

        // Explicit format with no extension route
        $noExtPath = $this->tempPath('') . '_no_extension';
        file_put_contents($noExtPath, "Col1,Col2\nVal1,Val2\n");

        $genNoExt = Excel::stream($noExtPath, ['format' => 'csv', 'start_row' => 2]);
        $rowsNoExt = iterator_to_array($genNoExt);
        $this->assertCount(1, $rowsNoExt);
        $this->assertSame('Val1', $rowsNoExt[2]['A']);
        @unlink($noExtPath);
    }
}
