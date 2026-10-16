<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\FormattedType;
use ExcelHelper\Types\RawType;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

final class RawAndFormattedTypeTest extends TestCase
{
    public function testRawTypeReadPreservesValue(): void
    {
        $type = new RawType();
        $ctx = new ColumnContext('raw_col');

        $this->assertSame(42, $type->read(42, $ctx));
        $this->assertSame('Hello', $type->read('Hello', $ctx));
        $this->assertSame(['nested'], $type->read(['nested'], $ctx));
        $this->assertNull($type->read(null, $ctx));
    }

    public function testRawTypeWrite(): void
    {
        $type = new RawType();
        $ctx = new ColumnContext('raw_col');

        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_BOOL, $type->write(true, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_NUMERIC, $type->write(123, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_STRING, $type->write('text', $ctx)->dataType);
    }

    public function testFormattedTypeRead(): void
    {
        $type = new FormattedType();
        $ctx = new ColumnContext('formatted_col');

        $this->assertSame('12.5%', $type->read('  12.5%  ', $ctx));
    }

    public function testFormattedTypeWriteThrows(): void
    {
        $type = new FormattedType();
        $ctx = new ColumnContext('formatted_col');

        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Type "formatted" is read-only and cannot be used for writing.');
        $type->write('12.5%', $ctx);
    }
}
