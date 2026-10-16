<?php

declare(strict_types=1);

namespace ExcelHelper\Contracts;

use ExcelHelper\Types\CellValue;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\TypeCastException;

/**
 * Two-way column data type interface (reading from cell and writing to cell).
 */
interface TypeInterface
{
    /**
     * Unique type identifier (e.g. 'string', 'int', 'date', 'money').
     */
    public function name(): string;

    /**
     * Convert raw cell value from spreadsheet into PHP type.
     *
     * @throws TypeCastException if the raw value cannot be cast to this type
     */
    public function read(mixed $value, ColumnContext $context): mixed;

    /**
     * Convert PHP value into a spreadsheet cell representation.
     */
    public function write(mixed $value, ColumnContext $context): CellValue;
}
