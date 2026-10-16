<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\InvalidConfigException;

/**
 * Formatted column type (extracts cell's formatted string display value; read-only).
 */
final class FormattedType implements TypeInterface
{
    public function name(): string
    {
        return 'formatted';
    }

    public function read(mixed $value, ColumnContext $context): string
    {
        $str = (string) $value;

        return $context->isTrim() ? trim($str) : $str;
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        throw InvalidConfigException::withMessage('Type "formatted" is read-only and cannot be used for writing.');
    }
}
