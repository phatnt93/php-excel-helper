<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Integration;

use ExcelHelper\Excel;
use ExcelHelper\Exceptions\SheetNotFoundException;
use ExcelHelper\Tests\TestCase;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

final class TemplateIntegrationTest extends TestCase
{
    public function testFillPlaceholdersAndDirectCells(): void
    {
        $templatePath = $this->tempPath('xlsx');

        // Create template
        $ss = new Spreadsheet();
        $sheet = $ss->getActiveSheet();
        $sheet->setTitle('Invoice');
        $sheet->setCellValue('A1', 'Report: {{report_name}}');
        $sheet->setCellValue('B2', '{{customer}}');
        $sheet->setCellValue('B3', '{{amount}}');
        $sheet->setCellValue('C5', 'Original Footer');

        $writer = IOFactory::createWriter($ss, IOFactory::WRITER_XLSX);
        $writer->save($templatePath);
        $ss->disconnectWorksheets();

        // Process template
        $outputPath = $this->tempPath('xlsx');
        Excel::template($templatePath)
            ->set('report_name', 'Sales Oct 2026')
            ->setValues([
                'customer' => 'Phat Nguyen',
                'amount' => 1250000.5,
                'D5' => 'Added Note',
            ])
            ->save($outputPath);

        // Verify output
        $res = IOFactory::load($outputPath);
        $ws = $res->getActiveSheet();

        $this->assertSame('Report: Sales Oct 2026', $ws->getCell('A1')->getValue());
        $this->assertSame('Phat Nguyen', $ws->getCell('B2')->getValue());
        $this->assertSame(1250000.5, $ws->getCell('B3')->getValue());
        $this->assertSame('Original Footer', $ws->getCell('C5')->getValue());
        $this->assertSame('Added Note', $ws->getCell('D5')->getValue());

        $res->disconnectWorksheets();
    }

    public function testInsertTableShiftsRowsDownward(): void
    {
        $templatePath = $this->tempPath('xlsx');

        // Create template with table placeholder and footer row
        $ss = new Spreadsheet();
        $sheet = $ss->getActiveSheet();
        $sheet->setCellValue('A1', 'Title');
        $sheet->setCellValue('A2', '{{item_name}}');
        $sheet->setCellValue('B2', '{{qty}}');
        $sheet->setCellValue('C2', '{{price}}');
        $sheet->setCellValue('A3', 'Footer Summary');
        $sheet->setCellValue('C3', '=SUM(C2:C2)');

        $writer = IOFactory::createWriter($ss, IOFactory::WRITER_XLSX);
        $writer->save($templatePath);
        $ss->disconnectWorksheets();

        $items = [
            ['item_name' => 'Laptop', 'qty' => 1, 'price' => 20000000],
            ['item_name' => 'Mouse', 'qty' => 2, 'price' => 500000],
            ['item_name' => 'Keyboard', 'qty' => 1, 'price' => 1200000],
        ];

        $outputPath = $this->tempPath('xlsx');
        Excel::template($templatePath)
            ->insertTable('item_name', $items)
            ->save($outputPath);

        $res = IOFactory::load($outputPath);
        $ws = $res->getActiveSheet();

        // Row 2: Laptop
        $this->assertSame('Laptop', $ws->getCell('A2')->getValue());
        $this->assertSame(1, $ws->getCell('B2')->getValue());
        $this->assertSame(20000000, $ws->getCell('C2')->getValue());

        // Row 3: Mouse
        $this->assertSame('Mouse', $ws->getCell('A3')->getValue());
        $this->assertSame(2, $ws->getCell('B3')->getValue());
        $this->assertSame(500000, $ws->getCell('C3')->getValue());

        // Row 4: Keyboard
        $this->assertSame('Keyboard', $ws->getCell('A4')->getValue());
        $this->assertSame(1, $ws->getCell('B4')->getValue());
        $this->assertSame(1200000, $ws->getCell('C4')->getValue());

        // Row 5: Footer Summary was pushed from row 3 down to row 5!
        $this->assertSame('Footer Summary', $ws->getCell('A5')->getValue());

        $res->disconnectWorksheets();
    }

    public function testTemplateInvalidSheetThrows(): void
    {
        $templatePath = $this->tempPath('xlsx');
        $ss = new Spreadsheet();
        IOFactory::createWriter($ss, IOFactory::WRITER_XLSX)->save($templatePath);
        $ss->disconnectWorksheets();

        $this->expectException(SheetNotFoundException::class);
        Excel::template($templatePath)->set('key', 'val', 'NonExistentSheet');
    }
    public function testInsertTableWithExplicitColumnsAndEmptyData(): void
    {
        $templatePath = $this->tempPath('xlsx');
        $ss = new Spreadsheet();
        $sheet = $ss->getActiveSheet();
        $sheet->setCellValue('A1', 'Header');
        $sheet->setCellValue('A2', '{{rows}}');
        $sheet->setCellValue('B2', '{{val}}');
        IOFactory::createWriter($ss, IOFactory::WRITER_XLSX)->save($templatePath);
        $ss->disconnectWorksheets();

        $tpl = Excel::template($templatePath);

        // Target not found returns self
        $tpl->insertTable('non_existent', [['a' => 1]]);

        // Empty data removes template row
        $tpl->insertTable('rows', []);

        $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Spreadsheet::class, $tpl->spreadsheet());
        $this->assertNotEmpty($tpl->toString());
    }

    public function testTemplateTypesBooleanNullFormula(): void
    {
        $templatePath = $this->tempPath('xlsx');
        $ss = new Spreadsheet();
        $sheet = $ss->getActiveSheet();
        $sheet->setCellValue('A1', '{{flag}}');
        $sheet->setCellValue('A2', '{{empty}}');
        $sheet->setCellValue('A3', '{{calc}}');
        IOFactory::createWriter($ss, IOFactory::WRITER_XLSX)->save($templatePath);
        $ss->disconnectWorksheets();

        $tpl = Excel::template($templatePath);
        $tpl->set('flag', true);
        $tpl->set('empty', null);
        $tpl->set('calc', '=1+1');

        $out = $this->tempPath('xlsx');
        $tpl->save($out);

        $loaded = IOFactory::load($out);
        $ws = $loaded->getActiveSheet();
        $this->assertTrue($ws->getCell('A1')->getValue());
        $this->assertNull($ws->getCell('A2')->getValue());
        $this->assertSame('=1+1', $ws->getCell('A3')->getValue());
        $loaded->disconnectWorksheets();
    }
}
