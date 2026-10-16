<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use DateTime;
use DateTimeImmutable;
use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\I18n\Translator;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\AutoType;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\DateTimeType;
use ExcelHelper\Types\DateType;
use ExcelHelper\Types\DecimalType;
use ExcelHelper\Types\FloatType;
use ExcelHelper\Types\FormulaType;
use ExcelHelper\Types\IntType;
use ExcelHelper\Types\LinkType;
use ExcelHelper\Types\MoneyType;
use ExcelHelper\Types\PercentType;
use ExcelHelper\Types\TimeType;
use ExcelHelper\Types\TypeCastException;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

final class TypesEdgeCasesTest extends TestCase
{
    public function testTypeCastExceptionGetters(): void
    {
        $e = new TypeCastException('invalid_date', ['value' => 'abc']);
        $this->assertSame('invalid_date', $e->getCodeString());
        $this->assertSame(['value' => 'abc'], $e->getParams());
    }

    public function testColumnContextGetters(): void
    {
        $ctx = new ColumnContext('col', ['trim' => false], [
            'date_system_1904' => true,
            'locale' => 'vi',
        ]);

        $this->assertFalse($ctx->isTrim());
        $this->assertTrue($ctx->isDateSystem1904());
        $this->assertSame('vi', $ctx->getLocale());
    }

    public function testTranslatorEdgeCases(): void
    {
        $this->assertSame(['en', 'vi'], Translator::availableLocales());
        $this->assertTrue(Translator::isLocaleSupported('EN'));
        $this->assertFalse(Translator::isLocaleSupported('fr'));

        $obj = new class {
            public function __toString(): string
            {
                return 'custom_object';
            }
        };

        $msg = Translator::trans('row.custom', ['message' => $obj]);
        $this->assertStringContainsString('custom_object', $msg);

        // Interpolate array value
        $msgArr = Translator::trans('row.custom', ['message' => ['nested' => 1]]);
        $this->assertStringContainsString('{"nested":1}', $msgArr);
    }

    public function testDateTypeEdgeCases(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('date_col');

        // Already DateTimeImmutable
        $immutable = new DateTimeImmutable('2026-10-16');
        $this->assertSame('2026-10-16', $type->read($immutable, $ctx));

        // Already DateTimeInterface (mutable)
        $mutable = new DateTime('2026-10-16');
        $this->assertSame('2026-10-16', $type->read($mutable, $ctx));

        // Empty string throws
        $this->expectException(TypeCastException::class);
        $type->read('', $ctx);
    }

    public function testDateTypeWriteWithTimestampAndUnsupported(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('date_col');

        // Timestamp
        $val = $type->write(1700000000, $ctx);
        $this->assertSame(DataType::TYPE_NUMERIC, $val->dataType);

        // String date
        $valStr = $type->write('2026-10-16', $ctx);
        $this->assertSame(DataType::TYPE_NUMERIC, $valStr->dataType);

        // Unparseable string date throws
        $this->expectException(WriterException::class);
        $type->write('not-a-date', $ctx);
    }

    public function testDateTypeWriteUnsupportedTypeThrows(): void
    {
        $type = new DateType();
        $ctx = new ColumnContext('date_col');

        $this->expectException(WriterException::class);
        $type->write(['array_not_allowed'], $ctx);
    }

    public function testDateTimeTypeEdgeCases(): void
    {
        $type = new DateTimeType();
        $ctx = new ColumnContext('datetime_col');

        $immutable = new DateTimeImmutable('2026-10-16 12:00:00');
        $this->assertSame('2026-10-16 12:00:00', $type->read($immutable, $ctx));

        $mutable = new DateTime('2026-10-16 12:00:00');
        $this->assertSame('2026-10-16 12:00:00', $type->read($mutable, $ctx));

        // Format null returns DateTimeImmutable
        $ctxNull = new ColumnContext('datetime_col', ['format' => null]);
        $res = $type->read('2026-10-16 12:00:00', $ctxNull);
        $this->assertInstanceOf(DateTimeImmutable::class, $res);

        // Write timestamp and null
        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_NUMERIC, $type->write(1700000000, $ctx)->dataType);

        // Write unparseable string throws
        $this->expectException(WriterException::class);
        $type->write('invalid-date', $ctx);
    }

    public function testDateTimeTypeWriteUnsupportedThrows(): void
    {
        $type = new DateTimeType();
        $ctx = new ColumnContext('datetime_col');

        $this->expectException(WriterException::class);
        $type->write(['unsupported'], $ctx);
    }

    public function testTimeTypeEdgeCases(): void
    {
        $type = new TimeType();
        $ctx = new ColumnContext('time_col');

        $immutable = new DateTimeImmutable('14:30:00');
        $this->assertSame('14:30:00', $type->read($immutable, $ctx));

        $mutable = new DateTime('14:30:00');
        $this->assertSame('14:30:00', $type->read($mutable, $ctx));

        $ctxNull = new ColumnContext('time_col', ['format' => null]);
        $res = $type->read('14:30:00', $ctxNull);
        $this->assertInstanceOf(DateTimeImmutable::class, $res);

        // Write timestamp and null
        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_NUMERIC, $type->write(1700000000, $ctx)->dataType);

        // Write unparseable string throws
        $this->expectException(WriterException::class);
        $type->write('invalid-time', $ctx);
    }

    public function testTimeTypeWriteUnsupportedThrows(): void
    {
        $type = new TimeType();
        $ctx = new ColumnContext('time_col');

        $this->expectException(WriterException::class);
        $type->write(['unsupported'], $ctx);
    }

    public function testLinkTypeEdgeCases(): void
    {
        $type = new LinkType();
        $ctx = new ColumnContext('link_col');

        $this->assertSame('https://example.com', $type->read('  https://example.com  ', $ctx));
        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);

        // Relative path
        $valRel = $type->write('/about-us', $ctx);
        $this->assertSame('/about-us', $valRel->hyperlink);

        $valRelDot = $type->write('./page.html', $ctx);
        $this->assertSame('./page.html', $valRelDot->hyperlink);

        // Invalid link format
        $this->expectException(WriterException::class);
        $type->write('invalid:scheme:test', $ctx);
    }

    public function testLinkTypeNonStringOrArrayThrows(): void
    {
        $type = new LinkType();
        $ctx = new ColumnContext('link_col');

        $this->expectException(WriterException::class);
        $type->write(12345, $ctx);
    }

    public function testPercentTypeEdgeCases(): void
    {
        $type = new PercentType();
        $ctx = new ColumnContext('discount');

        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);

        $this->expectException(TypeCastException::class);
        $type->read('invalid%', $ctx);
    }

    public function testPercentTypeWriteInvalidThrows(): void
    {
        $type = new PercentType();
        $ctx = new ColumnContext('discount');

        $this->expectException(WriterException::class);
        $type->write('not-a-number', $ctx);
    }

    public function testMoneyTypeEdgeCases(): void
    {
        $type = new MoneyType();
        $ctx = new ColumnContext('salary');

        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);

        $this->expectException(TypeCastException::class);
        $type->read('invalid-money', $ctx);
    }

    public function testMoneyTypeWriteInvalidThrows(): void
    {
        $type = new MoneyType();
        $ctx = new ColumnContext('salary');

        $this->expectException(WriterException::class);
        $type->write('not-a-number', $ctx);
    }

    public function testFloatTypeEmptyThrows(): void
    {
        $type = new FloatType();
        $ctx = new ColumnContext('f');

        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);

        $this->expectException(TypeCastException::class);
        $type->read('', $ctx);
    }

    public function testFloatTypeWriteInvalidThrows(): void
    {
        $type = new FloatType();
        $ctx = new ColumnContext('f');

        $this->expectException(WriterException::class);
        $type->write('not-number', $ctx);
    }

    public function testIntTypeEmptyThrows(): void
    {
        $type = new IntType();
        $ctx = new ColumnContext('i');

        $this->expectException(TypeCastException::class);
        $type->read('', $ctx);
    }

    public function testDecimalTypeEmptyThrows(): void
    {
        $type = new DecimalType();
        $ctx = new ColumnContext('d');

        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);

        $this->expectException(TypeCastException::class);
        $type->read('', $ctx);
    }

    public function testDecimalTypeWriteInvalidThrows(): void
    {
        $type = new DecimalType();
        $ctx = new ColumnContext('d');

        $this->expectException(WriterException::class);
        $type->write('not-number', $ctx);
    }

    public function testFormulaTypeReadAndWrite(): void
    {
        $type = new FormulaType();
        $ctx = new ColumnContext('calc');

        $this->assertSame('=A1+B1', $type->read('=A1+B1', $ctx));
        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);
    }

    public function testAutoTypeRead(): void
    {
        $type = new AutoType();
        $ctx = new ColumnContext('a');

        $this->assertSame('raw', $type->read('raw', $ctx));
    }
}
