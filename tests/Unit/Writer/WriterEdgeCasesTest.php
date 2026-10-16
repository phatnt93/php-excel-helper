<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Writer;

use ExcelHelper\Config\WriteConfigParser;
use ExcelHelper\Excel;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Writer\CellWriter;
use ExcelHelper\Writer\ExcelWriter;
use ExcelHelper\Writer\StyleMapper;
use ExcelHelper\Writer\WriteColumn;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

final class WriterEdgeCasesTest extends TestCase
{
    public function testExcelWriterProperties(): void
    {
        $wb = WriteConfigParser::parse([
            'properties' => [
                'creator' => 'John Doe',
                'title' => 'Doc Title',
                'subject' => 'Doc Subject',
                'description' => 'Doc Desc',
                'company' => 'Acme Corp',
                'category' => 'Reports',
                'keywords' => 'excel, test',
            ],
            'columns' => ['a' => 'A'],
            'data' => [['a' => 1]],
        ]);

        $spreadsheet = ExcelWriter::build($wb);
        $props = $spreadsheet->getProperties();

        $this->assertSame('John Doe', $props->getCreator());
        $this->assertSame('Doc Title', $props->getTitle());
        $this->assertSame('Doc Subject', $props->getSubject());
        $this->assertSame('Doc Desc', $props->getDescription());
        $this->assertSame('Acme Corp', $props->getCompany());
        $this->assertSame('Reports', $props->getCategory());
        $this->assertSame('excel, test', $props->getKeywords());

        $spreadsheet->disconnectWorksheets();
    }

    public function testExcelOutputSpreadsheetGetterAndCsvOptions(): void
    {
        $savePath = $this->tempPath('csv');

        $out = Excel::write([
            'format' => 'csv',
            'csv' => [
                'delimiter' => ';',
                'enclosure' => '"',
                'line_ending' => "\r\n",
                'bom' => false,
            ],
            'columns' => ['x' => 'Col X'],
            'data' => [['x' => 'Hello;World']],
        ]);

        $this->assertInstanceOf(Spreadsheet::class, $out->spreadsheet());
        $out->save($savePath);

        $content = (string) file_get_contents($savePath);
        $this->assertStringContainsString('Col X', $content);
        $this->assertStringContainsString('"Hello;World"', $content);
    }

    public function testCellWriterDefaultValueFallback(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $col = new WriteColumn(key: 'item', default: 'N/A');
        CellWriter::writeCell($sheet, 'A1', null, $col);

        $this->assertSame('N/A', $sheet->getCell('A1')->getValue());
        $spreadsheet->disconnectWorksheets();
    }

    public function testWriteConfigParserInvalidDataThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('"data" must be an iterable');

        WriteConfigParser::parse([
            'columns' => ['a' => 'A'],
            'data' => 'not-an-iterable',
        ]);
    }

    public function testWriteConfigParserInvalidTapThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('"tap" must be callable');

        WriteConfigParser::parse([
            'columns' => ['a' => 'A'],
            'tap' => 'not-callable',
        ]);
    }

    public function testStyleMapperAdditionalAlignments(): void
    {
        $res = StyleMapper::map([
            'strike' => true,
            'align' => 'justify',
            'valign' => 'justify',
        ]);

        $this->assertTrue($res['font']['strikethrough']);
        $this->assertSame('justify', $res['alignment']['horizontal']);
        $this->assertSame('justify', $res['alignment']['vertical']);
    }

    public function testSheetRendererRowHeight(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'row_height' => 35.0,
            'auto_size' => false,
            'columns' => ['val' => 'Value'],
            'data' => [['val' => 'Row 1']],
        ])->save($savePath);

        $this->assertFileExists($savePath);
    }
    public function testColumnValueCallback(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'columns' => [
                'val' => [
                    'header' => 'Value',
                    'value' => static fn($val, array $row): string => 'PREFIX_' . $val,
                ],
            ],
            'data' => [['val' => '123']],
        ])->save($savePath);

        $res = Excel::read($savePath, ['columns' => ['A' => 'v'], 'start_row' => 2]);
        $this->assertSame('PREFIX_123', $res->rows()[0]['v']);
    }

    public function testInvalidSheetInMultiSheetThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Each sheet configuration must be an array');

        WriteConfigParser::parse([
            'sheets' => ['not-an-array'],
        ]);
    }
    public function testBuildWithExistingSpreadsheet(): void
    {
        $existing = new Spreadsheet();
        $wb = WriteConfigParser::parse(['columns' => ['a' => 'A']]);
        $built = ExcelWriter::build($wb, $existing);
        $this->assertSame($existing, $built);
        $existing->disconnectWorksheets();
    }

    public function testSaveWithoutExtensionUsesDefinitionFormat(): void
    {
        $path = $this->tempPath('') . '_no_ext';
        $out = Excel::write(['format' => 'csv', 'columns' => ['a' => 'A']]);
        $saved = $out->save($path);
        $this->assertFileExists($saved);
        @unlink($saved);
    }
    public function testShowHeaderFalse(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'show_header' => false,
            'columns' => ['val' => 'My Header'],
            'data' => [['val' => 'First Row']],
        ])->save($savePath);

        $read = Excel::read($savePath, ['columns' => ['A' => 'v'], 'start_row' => 1]);
        $this->assertSame('First Row', $read->rows()[0]['v']);
    }

    public function testColumnAlignAndWrapOptions(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'columns' => [
                'val' => [
                    'header' => 'Value',
                    'align' => 'right',
                    'wrap' => true,
                ],
            ],
            'data' => [['val' => 'Line 1\nLine 2']],
        ])->save($savePath);

        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $style = $ss->getActiveSheet()->getStyle('A2');
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT, $style->getAlignment()->getHorizontal());
        $this->assertTrue($style->getAlignment()->getWrapText());
        $ss->disconnectWorksheets();
    }

    public function testWorkbookLocaleTotalsLabel(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'locale' => 'vi',
            'columns' => [
                'name' => 'Tên',
                'val' => ['header' => 'Số lượng', 'type' => 'number'],
            ],
            'totals' => ['val' => 'sum'],
            'data' => [['name' => 'Item 1', 'val' => 10]],
        ])->save($savePath);
        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $this->assertSame('Tổng cộng', $ss->getActiveSheet()->getCell('A3')->getValue());
        $ss->disconnectWorksheets();
    }
}
