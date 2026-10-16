<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Writer;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Writer\StyleMapper;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

final class StyleMapperTest extends TestCase
{
    public function testMapFontStyles(): void
    {
        $res = StyleMapper::map([
            'bold' => true,
            'italic' => true,
            'underline' => true,
            'strike' => true,
            'color' => '#FF0000',
            'size' => 14,
            'font' => 'Arial',
        ]);

        $this->assertTrue($res['font']['bold']);
        $this->assertTrue($res['font']['italic']);
        $this->assertTrue($res['font']['underline']);
        $this->assertTrue($res['font']['strikethrough']);
        $this->assertSame(['rgb' => 'FF0000'], $res['font']['color']);
        $this->assertSame(14.0, $res['font']['size']);
        $this->assertSame('Arial', $res['font']['name']);
    }

    public function testMapFill(): void
    {
        $res = StyleMapper::map(['background' => '#1F4E78']);

        $this->assertSame(Fill::FILL_SOLID, $res['fill']['fillType']);
        $this->assertSame(['rgb' => '1F4E78'], $res['fill']['startColor']);
    }

    public function testMapAlignment(): void
    {
        $res = StyleMapper::map([
            'align' => 'center',
            'valign' => 'middle',
            'wrap' => true,
        ]);

        $this->assertSame(Alignment::HORIZONTAL_CENTER, $res['alignment']['horizontal']);
        $this->assertSame(Alignment::VERTICAL_CENTER, $res['alignment']['vertical']);
        $this->assertTrue($res['alignment']['wrapText']);
    }

    public function testMapBorders(): void
    {
        $resBool = StyleMapper::map(['borders' => true]);
        $this->assertSame(Border::BORDER_THIN, $resBool['borders']['allBorders']['borderStyle']);

        $resString = StyleMapper::map(['borders' => 'double']);
        $this->assertSame(Border::BORDER_DOUBLE, $resString['borders']['allBorders']['borderStyle']);

        $resArray = StyleMapper::map([
            'borders' => [
                'top' => 'thin',
                'bottom' => 'thick',
                'color' => '#CCCCCC',
            ],
        ]);
        $this->assertSame(Border::BORDER_THIN, $resArray['borders']['top']['borderStyle']);
        $this->assertSame(Border::BORDER_THICK, $resArray['borders']['bottom']['borderStyle']);
        $this->assertSame(['rgb' => 'CCCCCC'], $resArray['borders']['top']['color']);
    }

    public function testUnknownStyleKeyThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Unknown configuration option "shadow"');

        StyleMapper::map(['shadow' => true]);
    }
}
