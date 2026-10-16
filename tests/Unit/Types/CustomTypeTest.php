<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Excel;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\CellValue;
use ExcelHelper\Types\ColumnContext;

final class CustomTypeTest extends TestCase
{
    public function testRegisterAndUseCustomMoneyVndType(): void
    {
        $moneyVndType = new class implements TypeInterface {
            public function name(): string
            {
                return 'money_vnd';
            }

            public function read(mixed $value, ColumnContext $context): float
            {
                if (is_numeric($value)) {
                    return (float) $value;
                }

                $clean = preg_replace('/[^\d]/u', '', (string) $value);

                return (float) $clean;
            }

            public function write(mixed $value, ColumnContext $context): CellValue
            {
                return CellValue::numeric((float) $value, '#,##0 "đ"');
            }
        };

        Excel::registerType($moneyVndType);

        $path = $this->tempPath('xlsx');

        // Write using custom type
        Excel::write([
            'columns' => [
                'service' => 'Dịch Vụ',
                'fee' => ['header' => 'Chi Phí', 'type' => 'money_vnd'],
            ],
            'data' => [
                ['service' => 'Bảo trì', 'fee' => 1500000],
            ],
        ])->save($path);

        // Read using custom type
        $readResult = Excel::read($path, [
            'header_row' => 1,
            'columns' => [
                ['header' => 'Dịch Vụ', 'key' => 'service'],
                ['header' => 'Chi Phí', 'key' => 'fee', 'type' => 'money_vnd'],
            ],
        ]);

        $this->assertSame(1, $readResult->rowCount());
        $this->assertSame('Bảo trì', $readResult->rows()[0]['service']);
        $this->assertSame(1500000.0, $readResult->rows()[0]['fee']);
    }
}
