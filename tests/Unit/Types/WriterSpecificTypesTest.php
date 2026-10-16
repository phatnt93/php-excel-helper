<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\AutoType;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\FormulaType;
use ExcelHelper\Types\LinkType;
use ExcelHelper\Types\MoneyType;
use ExcelHelper\Types\PercentType;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

final class WriterSpecificTypesTest extends TestCase
{
    public function testAutoTypeWrite(): void
    {
        $type = new AutoType();
        $ctx = new ColumnContext('auto_col');

        $this->assertSame(DataType::TYPE_NULL, $type->write(null, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_BOOL, $type->write(true, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_NUMERIC, $type->write(42, $ctx)->dataType);
        $this->assertSame(DataType::TYPE_NUMERIC, $type->write(42.5, $ctx)->dataType);

        // Security check: string starting with = is explicitly TYPE_STRING, not formula
        $val = $type->write('=SUM(A1:B1)', $ctx);
        $this->assertSame(DataType::TYPE_STRING, $val->dataType);
        $this->assertSame('=SUM(A1:B1)', $val->value);
    }

    public function testPercentType(): void
    {
        $type = new PercentType();
        $ctx = new ColumnContext('discount');

        $this->assertSame(0.15, $type->read('15%', $ctx));
        $this->assertSame(0.20, $type->read(0.20, $ctx));

        $val = $type->write(0.15, $ctx);
        $this->assertSame(0.15, $val->value);
        $this->assertSame('0.00%', $val->numberFormat);
    }

    public function testMoneyType(): void
    {
        $type = new MoneyType();
        $ctx = new ColumnContext('salary');

        $this->assertSame(1250000.0, $type->read('$1,250,000', $ctx));
        $this->assertSame(500000.0, $type->read('500.000 đ', new ColumnContext('salary', ['thousands_separator' => '.'])));

        $val = $type->write(1250000, $ctx);
        $this->assertSame(1250000.0, $val->value);
        $this->assertSame('#,##0', $val->numberFormat);
    }

    public function testFormulaType(): void
    {
        $type = new FormulaType();
        $ctx = new ColumnContext('total');

        $val1 = $type->write('=SUM(A1:A10)', $ctx);
        $this->assertSame(DataType::TYPE_FORMULA, $val1->dataType);
        $this->assertSame('=SUM(A1:A10)', $val1->value);

        $val2 = $type->write('SUM(A1:A10)', $ctx);
        $this->assertSame('=SUM(A1:A10)', $val2->value);

        $this->expectException(WriterException::class);
        $type->write('', $ctx);
    }

    public function testLinkType(): void
    {
        $type = new LinkType();
        $ctx = new ColumnContext('website');

        $val1 = $type->write('https://example.com', $ctx);
        $this->assertSame('https://example.com', $val1->hyperlink);
        $this->assertSame('https://example.com', $val1->value);

        $val2 = $type->write([
            'url' => 'https://example.com/docs',
            'label' => 'Documentation',
        ], $ctx);
        $this->assertSame('https://example.com/docs', $val2->hyperlink);
        $this->assertSame('Documentation', $val2->value);
    }

    public function testLinkTypeForbiddenSchemeThrows(): void
    {
        $type = new LinkType();
        $ctx = new ColumnContext('website');

        $this->expectException(WriterException::class);
        $this->expectExceptionMessage('Forbidden link scheme "javascript:"');
        $type->write('javascript:alert(1)', $ctx);
    }
}
