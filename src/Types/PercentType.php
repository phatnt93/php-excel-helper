<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;

/**
 * Percentage column type (0.15 = 15%).
 */
final class PercentType implements TypeInterface
{
    public function name(): string
    {
        return 'percent';
    }

    public function read(mixed $value, ColumnContext $context): float
    {
        if (is_int($value) || is_float($value)) {
            return (float) $value;
        }

        if (is_string($value)) {
            $str = $context->isTrim() ? trim($value) : $value;
            if (str_ends_with($str, '%')) {
                $numStr = rtrim($str, '%');
                if (is_numeric($numStr)) {
                    return (float) $numStr / 100.0;
                }
            } elseif (is_numeric($str)) {
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

        if (!is_numeric($value)) {
            throw WriterException::cannotWriteValue($context->key, sprintf('Value "%s" cannot be written as percent.', (string) $value));
        }

        /** @var string $formatCode */
        $formatCode = $context->get('format', '0.00%');

        return CellValue::numeric((float) $value, $formatCode);
    }
}
