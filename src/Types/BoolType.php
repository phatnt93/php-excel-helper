<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;

/**
 * Boolean column type.
 */
final class BoolType implements TypeInterface
{
    /**
     * @var list<string>
     */
    private const DEFAULT_TRUE_VALUES = ['true', '1', 'yes', 'y', 'on', 'đúng', 'có'];

    /**
     * @var list<string>
     */
    private const DEFAULT_FALSE_VALUES = ['false', '0', 'no', 'n', 'off', 'sai', 'không'];

    public function name(): string
    {
        return 'bool';
    }

    public function read(mixed $value, ColumnContext $context): bool
    {
        if (is_bool($value)) {
            return $value;
        }

        if (is_int($value) || is_float($value)) {
            if ($value === 1 || $value === 1.0) {
                return true;
            }
            if ($value === 0 || $value === 0.0) {
                return false;
            }

            throw new TypeCastException('invalid_bool', ['value' => $value]);
        }

        if (is_string($value)) {
            $str = mb_strtolower($context->isTrim() ? trim($value) : $value);

            if ($str === '') {
                throw new TypeCastException('invalid_bool', ['value' => $value]);
            }

            /** @var list<string> $trueValues */
            $trueValues = $context->get('true_values', self::DEFAULT_TRUE_VALUES);
            /** @var list<string> $falseValues */
            $falseValues = $context->get('false_values', self::DEFAULT_FALSE_VALUES);

            $normalizedTrue = array_map('mb_strtolower', $trueValues);
            $normalizedFalse = array_map('mb_strtolower', $falseValues);

            if (in_array($str, $normalizedTrue, true)) {
                return true;
            }

            if (in_array($str, $normalizedFalse, true)) {
                return false;
            }
        }

        throw new TypeCastException('invalid_bool', ['value' => $value]);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        /** @var ?array{0?: string, 1?: string, true?: string, false?: string} $labels */
        $labels = $context->get('labels');
        if (is_array($labels)) {
            $boolVal = (bool) $value;
            $label = $labels[$boolVal ? 'true' : 'false'] ?? $labels[$boolVal] ?? $labels[$boolVal ? 1 : 0] ?? null;
            if ($label !== null) {
                return CellValue::string($label);
            }
        }

        return CellValue::boolean((bool) $value);
    }
}
