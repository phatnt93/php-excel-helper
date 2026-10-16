<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\DecimalType;
use ExcelHelper\Types\FloatType;
use ExcelHelper\Types\IntType;
use ExcelHelper\Types\TypeCastException;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

final class NumericTypesTest extends TestCase
{
    public function testIntTypeReadValid(): void
    {
        $intType = new IntType();
        $ctx = new ColumnContext('age');

        $this->assertSame(42, $intType->read(42, $ctx));
        $this->assertSame(42, $intType->read('42', $ctx));
        $this->assertSame(-10, $intType->read('-10', $ctx));
        $this->assertSame(100, $intType->read(100.0, $ctx));
        $this->assertSame(1, $intType->read(true, $ctx));
        $this->assertSame(0, $intType->read(false, $ctx));

        // Thousands separator
        $ctxSep = new ColumnContext('amount', ['thousands_separator' => ',']);
        $this->assertSame(1234567, $intType->read('1,234,567', $ctxSep));
    }

    public function testIntTypeReadInvalidThrows(): void
    {
        $intType = new IntType();
        $ctx = new ColumnContext('age');

        $this->expectException(TypeCastException::class);
        $intType->read('abc', $ctx);
    }

    public function testIntTypeReadFractionalThrows(): void
    {
        $intType = new IntType();
        $ctx = new ColumnContext('age');

        $this->expectException(TypeCastException::class);
        $intType->read(12.34, $ctx);
    }

    public function testIntTypeWrite(): void
    {
        $intType = new IntType();
        $ctx = new ColumnContext('count');

        $val = $intType->write(42, $ctx);
        $this->assertSame(42, $val->value);
        $this->assertSame(DataType::TYPE_NUMERIC, $val->dataType);

        $nullVal = $intType->write(null, $ctx);
        $this->assertSame(DataType::TYPE_NULL, $nullVal->dataType);

        $this->expectException(WriterException::class);
        $intType->write('not_a_number', $ctx);
    }

    public function testFloatTypeReadValid(): void
    {
        $floatType = new FloatType();
        $ctx = new ColumnContext('rate');

        $this->assertSame(12.34, $floatType->read(12.34, $ctx));
        $this->assertSame(12.0, $floatType->read(12, $ctx));
        $this->assertSame(12.34, $floatType->read('12.34', $ctx));
        $this->assertSame(1.0, $floatType->read(true, $ctx));

        // Thousands and decimal separator
        $ctxSep = new ColumnContext('rate', [
            'thousands_separator' => '.',
            'decimal_separator' => ',',
        ]);
        $this->assertSame(1234.56, $floatType->read('1.234,56', $ctxSep));
    }

    public function testFloatTypeReadInvalidThrows(): void
    {
        $floatType = new FloatType();
        $ctx = new ColumnContext('rate');

        $this->expectException(TypeCastException::class);
        $floatType->read('not-numeric', $ctx);
    }

    public function testFloatTypeWrite(): void
    {
        $floatType = new FloatType();
        $ctx = new ColumnContext('rate');

        $val = $floatType->write(12.34, $ctx);
        $this->assertSame(12.34, $val->value);
        $this->assertSame(DataType::TYPE_NUMERIC, $val->dataType);
    }

    public function testDecimalTypeReadValid(): void
    {
        $decimalType = new DecimalType();
        $ctx = new ColumnContext('price', ['decimals' => 2]);

        $this->assertSame('12.50', $decimalType->read('12.5', $ctx));
        $this->assertSame('100.00', $decimalType->read(100, $ctx));
        $this->assertSame('99.99', $decimalType->read(99.99, $ctx));

        $ctxNoDecimals = new ColumnContext('price');
        $this->assertSame('123.456', $decimalType->read('123.456', $ctxNoDecimals));
    }

    public function testDecimalTypeReadInvalidThrows(): void
    {
        $decimalType = new DecimalType();
        $ctx = new ColumnContext('price');

        $this->expectException(TypeCastException::class);
        $decimalType->read('abc', $ctx);
    }

    public function testDecimalTypeWrite(): void
    {
        $decimalType = new DecimalType();
        $ctx = new ColumnContext('price', ['decimals' => 2]);

        $val = $decimalType->write('12.34', $ctx);
        $this->assertSame(12.34, $val->value);
        $this->assertSame('0.00', $val->numberFormat);
        $ctxZero = new ColumnContext('price', ['decimals' => 0]);
        $valZero = $decimalType->write('12.34', $ctxZero);
        $this->assertSame('0', $valZero->numberFormat);
    }
}
