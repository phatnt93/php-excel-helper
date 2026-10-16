<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\TypeRegistry;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Writes an individual cell using explicit PhpSpreadsheet data types and hyperlink handling.
 */
final class CellWriter
{
    /**
     * Write a single value to a worksheet cell according to column type configuration.
     *
     * @param array<string, mixed> $row
     * @return ?string The number format code to be applied, if specified by the type
     */
    public static function writeCell(
        Worksheet $worksheet,
        string $cellAddress,
        mixed $value,
        WriteColumn $column,
        array $row = [],
    ): ?string {
        if ($column->value !== null) {
            $value = ($column->value)($value, $row);
        } elseif ($value === null && $column->default !== null) {
            $value = $column->default;
        }

        $type = TypeRegistry::get($column->type);
        $colContext = new ColumnContext($column->key, $column->options);

        $cellValue = $type->write($value, $colContext);

        $worksheet->setCellValueExplicit($cellAddress, $cellValue->value, $cellValue->dataType);

        if ($cellValue->hyperlink !== null) {
            $worksheet->getCell($cellAddress)->getHyperlink()->setUrl($cellValue->hyperlink);
        }

        return $cellValue->numberFormat ?? $column->format;
    }
}
