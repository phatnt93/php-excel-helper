<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\BoolType;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\TypeCastException;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

final class BoolTypeTest extends TestCase
{
    private BoolType $type;

    protected function setUp(): void
    {
        parent::setUp();
        $this->type = new BoolType();
    }

    public function testReadNativeBooleans(): void
    {
        $ctx = new ColumnContext('active');
        $this->assertTrue($this->type->read(true, $ctx));
        $this->assertFalse($this->type->read(false, $ctx));
    }

    public function testReadIntegerNumbers(): void
    {
        $ctx = new ColumnContext('active');
        $this->assertTrue($this->type->read(1, $ctx));
        $this->assertTrue($this->type->read(1.0, $ctx));
        $this->assertFalse($this->type->read(0, $ctx));
        $this->assertFalse($this->type->read(0.0, $ctx));
    }

    public function testReadInvalidNumbersThrows(): void
    {
        $ctx = new ColumnContext('active');

        $this->expectException(TypeCastException::class);
        $this->type->read(2, $ctx);
    }

    public function testReadDefaultStringValues(): void
    {
        $ctx = new ColumnContext('active');

        // English
        $this->assertTrue($this->type->read('true', $ctx));
        $this->assertTrue($this->type->read('TRUE', $ctx));
        $this->assertTrue($this->type->read('yes', $ctx));
        $this->assertTrue($this->type->read('y', $ctx));
        $this->assertTrue($this->type->read('1', $ctx));
        $this->assertTrue($this->type->read('on', $ctx));

        // Vietnamese
        $this->assertTrue($this->type->read('đúng', $ctx));
        $this->assertTrue($this->type->read('ĐÚNG', $ctx));
        $this->assertTrue($this->type->read('có', $ctx));

        // False values
        $this->assertFalse($this->type->read('false', $ctx));
        $this->assertFalse($this->type->read('no', $ctx));
        $this->assertFalse($this->type->read('n', $ctx));
        $this->assertFalse($this->type->read('0', $ctx));
        $this->assertFalse($this->type->read('off', $ctx));
        $this->assertFalse($this->type->read('sai', $ctx));
        $this->assertFalse($this->type->read('không', $ctx));
    }

    public function testReadCustomValues(): void
    {
        $ctx = new ColumnContext('is_vip', [
            'true_values' => ['vip', 'gold'],
            'false_values' => ['standard', 'regular'],
        ]);

        $this->assertTrue($this->type->read('vip', $ctx));
        $this->assertFalse($this->type->read('standard', $ctx));
    }

    public function testReadInvalidStringThrows(): void
    {
        $ctx = new ColumnContext('active');

        $this->expectException(TypeCastException::class);
        $this->type->read('maybe', $ctx);
    }

    public function testWrite(): void
    {
        $ctx = new ColumnContext('active');

        $val = $this->type->write(true, $ctx);
        $this->assertTrue($val->value);
        $this->assertSame(DataType::TYPE_BOOL, $val->dataType);

        $nullVal = $this->type->write(null, $ctx);
        $this->assertSame(DataType::TYPE_NULL, $nullVal->dataType);

        // Custom labels
        $ctxLabels = new ColumnContext('active', [
            'labels' => [true => 'Active', false => 'Inactive'],
        ]);
        $valLabeled = $this->type->write(true, $ctxLabels);
        $this->assertSame('Active', $valLabeled->value);
        $this->assertSame(DataType::TYPE_STRING, $valLabeled->dataType);
    }
}
