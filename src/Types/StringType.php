<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use DateTimeInterface;
use ExcelHelper\Contracts\TypeInterface;

/**
 * String column type.
 */
final class StringType implements TypeInterface
{
    public function name(): string
    {
        return 'string';
    }

    public function read(mixed $value, ColumnContext $context): string
    {
        if ($value === null) {
            return '';
        }

        if (is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }

        if ($value instanceof DateTimeInterface) {
            $format = $context->getDateFormat() ?? 'Y-m-d';

            return $value->format($format);
        }

        if (is_int($value)) {
            return (string) $value;
        }

        if (is_float($value)) {
            // Avoid scientific notation for integer-like floats (e.g. 1.0E+15)
            if (floor($value) === $value && abs($value) < 1.0E+16) {
                return sprintf('%.0f', $value);
            }

            return rtrim(rtrim(sprintf('%.10F', $value), '0'), '.');
        }

        $str = (string) $value;

        return $context->isTrim() ? trim($str) : $str;
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        if (is_bool($value)) {
            return CellValue::string($value ? 'TRUE' : 'FALSE');
        }

        if ($value instanceof DateTimeInterface) {
            $format = $context->getDateFormat() ?? 'Y-m-d';

            return CellValue::string($value->format($format));
        }

        return CellValue::string((string) $value);
    }
}
