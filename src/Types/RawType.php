<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

/**
 * Raw column type (preserves whatever value and type PhpSpreadsheet returns without casting).
 */
final class RawType implements TypeInterface
{
    public function name(): string
    {
        return 'raw';
    }

    public function read(mixed $value, ColumnContext $context): mixed
    {
        return $value;
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        if (is_bool($value)) {
            return CellValue::boolean($value);
        }

        if (is_int($value) || is_float($value)) {
            return CellValue::numeric($value);
        }

        return new CellValue($value, DataType::TYPE_STRING);
    }
}
