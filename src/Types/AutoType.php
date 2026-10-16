<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

/**
 * Auto column type (infers scalar types safely; strings are explicitly stored as text to prevent formula injection).
 */
final class AutoType implements TypeInterface
{
    public function name(): string
    {
        return 'auto';
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

        // Strings are strictly TYPE_STRING, never evaluated as formulas
        return new CellValue((string) $value, DataType::TYPE_STRING);
    }
}
