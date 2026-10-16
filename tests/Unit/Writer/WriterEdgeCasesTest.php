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
    public function testImageInsertion(): void
    {
        $imgPath = $this->tempPath('png');
        $img = imagecreatetruecolor(20, 20);
        if ($img !== false) {
            imagepng($img, $imgPath);
            imagedestroy($img);
        }

        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'images' => [
                [
                    'path' => $imgPath,
                    'cell' => 'B2',
                    'width' => 120,
                    'height' => 60,
                    'name' => 'Logo',
                ],
            ],
            'columns' => ['name' => 'Name'],
            'data' => [['name' => 'Item 1']],
        ])->save($savePath);

        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $drawings = $ss->getActiveSheet()->getDrawingCollection();
        $this->assertCount(1, $drawings);
        $this->assertSame('B2', $drawings[0]->getCoordinates());
        $this->assertSame(120, $drawings[0]->getWidth());
        $this->assertSame(60, $drawings[0]->getHeight());
        $this->assertSame('Logo', $drawings[0]->getName());
        $ss->disconnectWorksheets();
    }
    public function testColumnDataValidation(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'columns' => [
                'status' => [
                    'header' => 'Trạng thái',
                    'validation' => ['Chờ duyệt', 'Đã duyệt', 'Từ chối'],
                ],
                'age' => [
                    'header' => 'Tuổi',
                    'validation' => [
                        'type' => 'integer',
                        'min' => 18,
                        'max' => 60,
                    ],
                ],
            ],
            'data' => [
                ['status' => 'Chờ duyệt', 'age' => 25],
            ],
        ])->save($savePath);

        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $sheet = $ss->getActiveSheet();

        $valStatus = $sheet->getCell('A2')->getDataValidation();
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_LIST, $valStatus->getType());
        $this->assertSame('"Chờ duyệt,Đã duyệt,Từ chối"', $valStatus->getFormula1());
        $this->assertTrue($valStatus->getShowDropDown());

        $valAge = $sheet->getCell('B2')->getDataValidation();
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_WHOLE, $valAge->getType());
        $this->assertSame('18', $valAge->getFormula1());
        $this->assertSame('60', $valAge->getFormula2());

        $ss->disconnectWorksheets();
    }
    public function testConditionalFormatting(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'columns' => [
                'val' => [
                    'header' => 'Value',
                    'conditional' => [
                        [
                            'operator' => '<',
                            'value' => 0,
                            'color' => 'FF0000',
                            'bold' => true,
                        ],
                    ],
                ],
            ],
            'data' => [
                ['val' => -50],
                ['val' => 100],
            ],
        ])->save($savePath);

        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $sheet = $ss->getActiveSheet();

        $condStyles = $sheet->getStyle('A2:A3')->getConditionalStyles();
        $this->assertCount(1, $condStyles);
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_LESSTHAN, $condStyles[0]->getOperatorType());
        $this->assertSame(['0'], $condStyles[0]->getConditions());
        $this->assertSame('FF0000', $condStyles[0]->getStyle()->getFont()->getColor()->getRGB());
        $this->assertTrue($condStyles[0]->getStyle()->getFont()->getBold());

        $ss->disconnectWorksheets();
    }

    public function testDataValidationDecimalAndTextLength(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'columns' => [
                'rating' => [
                    'header' => 'Rating',
                    'validation' => [
                        'type' => 'decimal',
                        'min' => 1.0,
                        'max' => 5.0,
                    ],
                ],
                'comment' => [
                    'header' => 'Comment',
                    'validation' => [
                        'type' => 'text_length',
                        'min' => 5,
                        'max' => 100,
                    ],
                ],
            ],
            'data' => [
                ['rating' => 4.5, 'comment' => 'Great product!'],
            ],
        ])->save($savePath);

        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $sheet = $ss->getActiveSheet();

        $valRating = $sheet->getCell('A2')->getDataValidation();
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_DECIMAL, $valRating->getType());
        $this->assertSame('1', $valRating->getFormula1());
        $this->assertSame('5', $valRating->getFormula2());

        $valComment = $sheet->getCell('B2')->getDataValidation();
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_TEXTLENGTH, $valComment->getType());
        $this->assertSame('5', $valComment->getFormula1());
        $this->assertSame('100', $valComment->getFormula2());

        $ss->disconnectWorksheets();
    }

    public function testConditionalFormattingSheetLevelAndOperators(): void
    {
        $savePath = $this->tempPath('xlsx');

        Excel::write([
            'conditional_formatting' => [
                [
                    'range' => 'B2:B10',
                    'operator' => 'between',
                    'value' => [10, 50],
                    'bg' => 'FFFF00',
                    'italic' => true,
                ],
            ],
            'columns' => [
                'val' => [
                    'header' => 'Val',
                    'conditional' => [
                        ['operator' => '>', 'value' => 100, 'color' => '00FF00'],
                        ['operator' => '<=', 'value' => 50, 'color' => '0000FF'],
                        ['operator' => '==', 'value' => 75, 'color' => 'FF00FF'],
                        ['operator' => '!=', 'value' => 0, 'color' => '123456'],
                    ],
                ],
                'num' => 'Num',
            ],
            'data' => [
                ['val' => 75, 'num' => 20],
            ],
        ])->save($savePath);

        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $sheet = $ss->getActiveSheet();

        $sheetStyles = $sheet->getStyle('B2:B10')->getConditionalStyles();
        $this->assertCount(1, $sheetStyles);
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_BETWEEN, $sheetStyles[0]->getOperatorType());
        $this->assertSame(['10', '50'], $sheetStyles[0]->getConditions());
        $this->assertTrue($sheetStyles[0]->getStyle()->getFont()->getItalic());

        $colStyles = $sheet->getStyle('A2:A2')->getConditionalStyles();
        $this->assertCount(4, $colStyles);
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHAN, $colStyles[0]->getOperatorType());
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_LESSTHANOREQUAL, $colStyles[1]->getOperatorType());
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_EQUAL, $colStyles[2]->getOperatorType());
        $this->assertSame(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_NOTEQUAL, $colStyles[3]->getOperatorType());

        $ss->disconnectWorksheets();
    }

    public function testWriterBuilderFluentImagesAndMethods(): void
    {
        $imgPath = $this->tempPath('png');
        $img = imagecreatetruecolor(10, 10);
        if ($img !== false) {
            imagepng($img, $imgPath);
            imagedestroy($img);
        }

        $savePath = $this->tempPath('xlsx');

        Excel::writer()
            ->locale('vi')
            ->showHeader(true)
            ->image($imgPath, 'A1', 50, 50)
            ->images([['path' => $imgPath, 'cell' => 'C1']])
            ->column('name', 'Tên')
            ->data([['name' => 'Test']])
            ->save($savePath);

        $ss = \PhpOffice\PhpSpreadsheet\IOFactory::load($savePath);
        $drawings = $ss->getActiveSheet()->getDrawingCollection();
        $this->assertCount(2, $drawings);
        $ss->disconnectWorksheets();
    }
}
