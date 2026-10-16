<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;

/**
 * Decimal column type (returns formatted string to preserve decimal precision without floating-point artifacts).
 */
final class DecimalType implements TypeInterface
{
    public function name(): string
    {
        return 'decimal';
    }

    public function read(mixed $value, ColumnContext $context): string
    {
        if (is_int($value)) {
            $decimals = $context->get('decimals');
            if (is_int($decimals) && $decimals >= 0) {
                return number_format((float) $value, $decimals, '.', '');
            }

            return (string) $value;
        }

        if (is_float($value)) {
            $decimals = $context->get('decimals');
            if (is_int($decimals) && $decimals >= 0) {
                return number_format($value, $decimals, '.', '');
            }

            return rtrim(rtrim(sprintf('%.10F', $value), '0'), '.');
        }

        if (is_string($value)) {
            $str = $context->isTrim() ? trim($value) : $value;

            if ($str === '') {
                throw new TypeCastException('invalid_decimal', ['value' => $value]);
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
                $decimals = $context->get('decimals');
                if (is_int($decimals) && $decimals >= 0) {
                    return number_format((float) $str, $decimals, '.', '');
                }

                return $str;
            }
        }

        throw new TypeCastException('invalid_decimal', ['value' => $value]);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        if (!is_numeric($value)) {
            throw WriterException::cannotWriteValue($context->key, sprintf('Value "%s" cannot be written as decimal.', (string) $value));
        }

        $decimals = $context->get('decimals');
        if (is_int($decimals)) {
            $formatCode = $decimals > 0 ? '0.' . str_repeat('0', $decimals) : '0';
        } else {
            $formatCode = '0.00';
        }

        return CellValue::numeric((float) $value, $formatCode);
    }
}
