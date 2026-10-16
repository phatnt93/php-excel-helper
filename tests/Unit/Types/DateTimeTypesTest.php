<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use DateTime;
use DateTimeImmutable;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\DateTimeType;
use ExcelHelper\Types\DateType;
use ExcelHelper\Types\TimeType;
use ExcelHelper\Types\TypeCastException;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Shared\Date as PhpSpreadsheetDate;

final class DateTimeTypesTest extends TestCase
{
    public function testDateTypeReadExcelSerial(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('birthday');

        // Excel serial for 2026-10-16
        $serial = PhpSpreadsheetDate::stringToExcel('2026-10-16');
        $this->assertSame('2026-10-16', $type->read($serial, $ctx));
    }

    public function testDateTypeRead1904Calendar(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('birthday', [], ['date_system_1904' => true]);

        // In 1904 calendar, day 0 is Jan 1, 1904
        $res = $type->read(0, $ctx);
        $this->assertSame('1904-01-01', $res);
    }

    public function testDateTypeReadStrings(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('birthday');

        $this->assertSame('2026-10-16', $type->read('2026-10-16', $ctx));
        $this->assertSame('2026-10-16', $type->read('16/10/2026', $ctx));
    }

    public function testDateTypeReadInvalidDateRejectsOverflow(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('birthday', ['input_format' => 'd/m/Y']);

        // Feb 31 does not exist
        $this->expectException(TypeCastException::class);
        $type->read('31/02/2026', $ctx);
    }

    public function testDateTypeReadReturnsImmutableWhenFormatNull(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('birthday', ['format' => null]);

        $res = $type->read('2026-10-16', $ctx);
        $this->assertInstanceOf(DateTimeImmutable::class, $res);
        $this->assertSame('2026-10-16', $res->format('Y-m-d'));
    }

    public function testDateTypeWrite(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('birthday');

        $dt = new DateTime('2026-10-16');
        $cellVal = $type->write($dt, $ctx);

        $this->assertSame(DataType::TYPE_NUMERIC, $cellVal->dataType);
        $this->assertSame('yyyy-mm-dd', $cellVal->numberFormat);
    }

    public function testDateTimeTypeReadAndWrite(): void
    {
        $type = new DateTimeType();
        $ctx = new ColumnContext('logged_at');

        $this->assertSame('2026-10-16 14:30:00', $type->read('2026-10-16 14:30:00', $ctx));

        $val = $type->write('2026-10-16 14:30:00', $ctx);
        $this->assertSame(DataType::TYPE_NUMERIC, $val->dataType);
        $this->assertSame('yyyy-mm-dd hh:mm:ss', $val->numberFormat);
    }

    public function testTimeTypeReadAndWrite(): void
    {
        $type = new TimeType();
        $ctx = new ColumnContext('start_time');

        $this->assertSame('14:30:00', $type->read('14:30:00', $ctx));
        $this->assertSame('14:30:00', $type->read('14:30', $ctx));

        // 0.5 in Excel serial is 12:00:00 noon
        $this->assertSame('12:00:00', $type->read(0.5, $ctx));

        $val = $type->write('14:30:00', $ctx);
        $this->assertSame(DataType::TYPE_NUMERIC, $val->dataType);
        $this->assertSame('hh:mm:ss', $val->numberFormat);
    }

    public function testTimeTypeReadInvalidThrows(): void
    {
        $type = new TimeType();
        $ctx = new ColumnContext('start_time');

        $this->expectException(TypeCastException::class);
        $type->read('invalid-time-string', $ctx);
    }
}
