<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use DateTimeImmutable;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\StringType;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

final class StringTypeTest extends TestCase
{
    private StringType $type;

    protected function setUp(): void
    {
        parent::setUp();
        $this->type = new StringType();
    }

    public function testName(): void
    {
        $this->assertSame('string', $this->type->name());
    }

    public function testReadBasicStrings(): void
    {
        $ctx = new ColumnContext('title');
        $this->assertSame('Hello', $this->type->read('Hello', $ctx));
        $this->assertSame('', $this->type->read(null, $ctx));
        $this->assertSame('', $this->type->read('', $ctx));
    }

    public function testReadTrimming(): void
    {
        $ctxTrim = new ColumnContext('title', [], ['trim' => true]);
        $this->assertSame('spaced', $this->type->read('  spaced  ', $ctxTrim));

        $ctxNoTrim = new ColumnContext('title', [], ['trim' => false]);
        $this->assertSame('  spaced  ', $this->type->read('  spaced  ', $ctxNoTrim));
    }

    public function testReadBooleans(): void
    {
        $ctx = new ColumnContext('status');
        $this->assertSame('TRUE', $this->type->read(true, $ctx));
        $this->assertSame('FALSE', $this->type->read(false, $ctx));
    }

    public function testReadNumbers(): void
    {
        $ctx = new ColumnContext('code');
        $this->assertSame('123', $this->type->read(123, $ctx));
        $this->assertSame('123.45', $this->type->read(123.45, $ctx));

        // Float representing large integer without exponent notation
        $this->assertSame('1000000000000000', $this->type->read(1.0E+15, $ctx));
    }

    public function testReadDateTimeInterface(): void
    {
        $ctx = new ColumnContext('created_at', ['format' => 'd/m/Y']);
        $dt = new DateTimeImmutable('2026-10-16 08:30:00');
        $this->assertSame('16/10/2026', $this->type->read($dt, $ctx));
    }

    public function testWrite(): void
    {
        $ctx = new ColumnContext('title');
        $cellVal = $this->type->write('Hello', $ctx);

        $this->assertSame('Hello', $cellVal->value);
        $this->assertSame(DataType::TYPE_STRING, $cellVal->dataType);

        $nullVal = $this->type->write(null, $ctx);
        $this->assertSame(DataType::TYPE_NULL, $nullVal->dataType);

        $boolVal = $this->type->write(true, $ctx);
        $this->assertSame('TRUE', $boolVal->value);

        $dtVal = $this->type->write(new DateTimeImmutable('2026-10-16'), $ctx);
        $this->assertSame('2026-10-16', $dtVal->value);
    }
}
