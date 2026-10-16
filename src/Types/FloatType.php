<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;

/**
 * Float / Number column type.
 */
final class FloatType implements TypeInterface
{
    public function name(): string
    {
        return 'float';
    }

    public function read(mixed $value, ColumnContext $context): float
    {
        if (is_float($value)) {
            return $value;
        }

        if (is_int($value)) {
            return (float) $value;
        }

        if (is_bool($value)) {
            return $value ? 1.0 : 0.0;
        }

        if (is_string($value)) {
            $str = $context->isTrim() ? trim($value) : $value;

            if ($str === '') {
                throw new TypeCastException('invalid_float', ['value' => $value]);
            }

            /** @var ?string $thousandsSep */
            $thousandsSep = $context->get('thousands_separator');
            if ($thousandsSep !== null && $thousandsSep !== '') {
                $str = str_replace($thousandsSep, '', $str);
            }

            /** @var ?string $decimalSep */
            $decimalSep = $context->get('decimal_separator');
            if ($decimalSep !== null && $decimalSep !== '.' && $decimalSep !== '') {
                $str = str_replace($decimalSep, '.', $str);
            }

            if (is_numeric($str)) {
                return (float) $str;
            }
        }

        throw new TypeCastException('invalid_float', ['value' => $value]);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        if (!is_numeric($value) && !is_bool($value)) {
            throw WriterException::cannotWriteValue($context->key, sprintf('Value "%s" cannot be written as float.', (string) $value));
        }

        return CellValue::numeric((float) $value);
    }
}
