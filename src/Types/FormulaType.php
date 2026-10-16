<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;

/**
 * Formula column type.
 */
final class FormulaType implements TypeInterface
{
    public function name(): string
    {
        return 'formula';
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

        if (!is_string($value) || trim($value) === '') {
            throw WriterException::cannotWriteValue($context->key, 'Formula value must be a non-empty string.');
        }

        $formula = trim($value);
        if (!str_starts_with($formula, '=')) {
            $formula = '=' . $formula;
        }

        return CellValue::formula($formula);
    }
}
