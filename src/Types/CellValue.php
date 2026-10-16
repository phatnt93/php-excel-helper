<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use PhpOffice\PhpSpreadsheet\Cell\DataType;

/**
 * Value object representing a cell value, its explicit data type, number format, and optional hyperlink.
 */
final class CellValue
{
    public function __construct(
        public readonly mixed $value,
        public readonly string $dataType = DataType::TYPE_STRING,
        public readonly ?string $numberFormat = null,
        public readonly ?string $hyperlink = null,
    ) {}

    public static function string(string $value): self
    {
        return new self($value, DataType::TYPE_STRING);
    }

    public static function numeric(int|float $value, ?string $numberFormat = null): self
    {
        return new self($value, DataType::TYPE_NUMERIC, $numberFormat);
    }

    public static function boolean(bool $value): self
    {
        return new self($value, DataType::TYPE_BOOL);
    }

    public static function formula(string $formula): self
    {
        return new self($formula, DataType::TYPE_FORMULA);
    }

    public static function null(): self
    {
        return new self(null, DataType::TYPE_NULL);
    }
}
