<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Support;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Support\Coordinate;
use ExcelHelper\Tests\TestCase;

final class CoordinateTest extends TestCase
{
    public function testIsColumnLetter(): void
    {
        $this->assertTrue(Coordinate::isColumnLetter('A'));
        $this->assertTrue(Coordinate::isColumnLetter('z'));
        $this->assertTrue(Coordinate::isColumnLetter('AA'));
        $this->assertTrue(Coordinate::isColumnLetter('XFD'));

        $this->assertFalse(Coordinate::isColumnLetter('XFE'));
        $this->assertFalse(Coordinate::isColumnLetter('AAAA'));
        $this->assertFalse(Coordinate::isColumnLetter(''));
        $this->assertFalse(Coordinate::isColumnLetter('1'));
        $this->assertFalse(Coordinate::isColumnLetter('A1'));
        $this->assertFalse(Coordinate::isColumnLetter(' '));
    }

    public function testColumnIndexFromInteger(): void
    {
        $this->assertSame(1, Coordinate::columnIndex(1));
        $this->assertSame(26, Coordinate::columnIndex(26));
        $this->assertSame(27, Coordinate::columnIndex(27));
        $this->assertSame(16384, Coordinate::columnIndex(16384));
    }

    public function testColumnIndexFromString(): void
    {
        $this->assertSame(1, Coordinate::columnIndex('A'));
        $this->assertSame(1, Coordinate::columnIndex('a'));
        $this->assertSame(2, Coordinate::columnIndex('B'));
        $this->assertSame(26, Coordinate::columnIndex('Z'));
        $this->assertSame(27, Coordinate::columnIndex('AA'));
        $this->assertSame(16384, Coordinate::columnIndex('XFD'));

        // Numeric strings
        $this->assertSame(1, Coordinate::columnIndex('1'));
        $this->assertSame(26, Coordinate::columnIndex('26'));
    }

    public function testColumnIndexThrowsOnOutOfBoundsInt(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Column index must be between 1 and 16384');
        Coordinate::columnIndex(0);
    }

    public function testColumnIndexThrowsOnTooHighInt(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Column index must be between 1 and 16384');
        Coordinate::columnIndex(16385);
    }

    public function testColumnIndexThrowsOnInvalidString(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Invalid column identifier "XFE"');
        Coordinate::columnIndex('XFE');
    }

    public function testColumnLetterFromInteger(): void
    {
        $this->assertSame('A', Coordinate::columnLetter(1));
        $this->assertSame('Z', Coordinate::columnLetter(26));
        $this->assertSame('AA', Coordinate::columnLetter(27));
        $this->assertSame('XFD', Coordinate::columnLetter(16384));
    }

    public function testColumnLetterFromString(): void
    {
        $this->assertSame('A', Coordinate::columnLetter('A'));
        $this->assertSame('B', Coordinate::columnLetter('b'));
        $this->assertSame('A', Coordinate::columnLetter('1'));
        $this->assertSame('AA', Coordinate::columnLetter('27'));
    }

    public function testColumnLetterThrowsOnOutOfBounds(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Column index must be between 1 and 16384');
        Coordinate::columnLetter(0);
    }

    public function testColumnLetterThrowsOnInvalidString(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Invalid column identifier "INVALID"');
        Coordinate::columnLetter('INVALID');
    }

    public function testParseCell(): void
    {
        $this->assertSame(['col' => 1, 'row' => 1], Coordinate::parseCell('A1'));
        $this->assertSame(['col' => 2, 'row' => 3], Coordinate::parseCell('B3'));
        $this->assertSame(['col' => 2, 'row' => 3], Coordinate::parseCell('$B$3'));
        $this->assertSame(['col' => 27, 'row' => 100], Coordinate::parseCell('AA100'));
    }

    public function testParseCellThrowsOnInvalidCell(): void
    {
        $this->expectException(InvalidConfigException::class);
        Coordinate::parseCell('INVALID_CELL');
    }

    public function testFormatCell(): void
    {
        $this->assertSame('A1', Coordinate::formatCell(1, 1));
        $this->assertSame('B3', Coordinate::formatCell(2, 3));
        $this->assertSame('AA10', Coordinate::formatCell(27, 10));
    }

    public function testFormatCellThrowsOnInvalidRow(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Row index must be >= 1');
        Coordinate::formatCell(1, 0);
    }
}
