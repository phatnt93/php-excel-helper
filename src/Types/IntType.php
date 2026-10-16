<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;

/**
 * Integer column type.
 */
final class IntType implements TypeInterface
{
    public function name(): string
    {
        return 'int';
    }

    public function read(mixed $value, ColumnContext $context): int
    {
        if (is_int($value)) {
            return $value;
        }

        if (is_bool($value)) {
            return $value ? 1 : 0;
        }

        if (is_float($value)) {
            if (floor($value) === $value) {
                return (int) $value;
            }

            throw new TypeCastException('invalid_int', ['value' => $value]);
        }

        if (is_string($value)) {
            $str = $context->isTrim() ? trim($value) : $value;

            if ($str === '') {
                throw new TypeCastException('invalid_int', ['value' => $value]);
            }

            /** @var ?string $thousandsSep */
            $thousandsSep = $context->get('thousands_separator');
            if ($thousandsSep !== null && $thousandsSep !== '') {
                $str = str_replace($thousandsSep, '', $str);
            }

            if (preg_match('/^[+-]?\d+$/D', $str)) {
                return (int) $str;
            }

            if (is_numeric($str)) {
                $floatVal = (float) $str;
                if (floor($floatVal) === $floatVal) {
                    return (int) $floatVal;
                }
            }
        }

        throw new TypeCastException('invalid_int', ['value' => $value]);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        if (!is_numeric($value) && !is_bool($value)) {
            throw WriterException::cannotWriteValue($context->key, sprintf('Value "%s" cannot be written as integer.', (string) $value));
        }

        return CellValue::numeric((int) $value);
    }
}
