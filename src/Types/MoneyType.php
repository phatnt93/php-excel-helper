<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;

/**
 * Money column type.
 */
final class MoneyType implements TypeInterface
{
    public function name(): string
    {
        return 'money';
    }

    public function read(mixed $value, ColumnContext $context): float
    {
        if (is_int($value) || is_float($value)) {
            return (float) $value;
        }

        if (is_string($value)) {
            $str = $context->isTrim() ? trim($value) : $value;
            // Clean out common currency symbols and thousands spaces/commas
            $cleaned = preg_replace('/[^\d.,+-]/u', '', $str);
            if ($cleaned !== null && $cleaned !== '') {
                /** @var ?string $thousandsSep */
                $thousandsSep = $context->get('thousands_separator', ',');
                if ($thousandsSep !== null && $thousandsSep !== '') {
                    $cleaned = str_replace($thousandsSep, '', $cleaned);
                }

                /** @var ?string $decimalSep */
                $decimalSep = $context->get('decimal_separator');
                if ($decimalSep !== null && $decimalSep !== '.' && $decimalSep !== '') {
                    $cleaned = str_replace($decimalSep, '.', $cleaned);
                }

                if (is_numeric($cleaned)) {
                    return (float) $cleaned;
                }
            }
        }

        throw new TypeCastException('invalid_float', ['value' => $value]);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        if (!is_numeric($value)) {
            throw WriterException::cannotWriteValue($context->key, sprintf('Value "%s" cannot be written as money.', (string) $value));
        }

        /** @var string $formatCode */
        $formatCode = $context->get('format', '#,##0');

        return CellValue::numeric((float) $value, $formatCode);
    }
}
