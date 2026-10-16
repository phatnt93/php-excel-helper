<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

/**
 * Metadata about the cell ranges rendered on a worksheet, passed to tap() callbacks.
 */
final class SheetLayout
{
    public function __construct(
        public readonly ?string $titleCell,
        public readonly ?int $headerRow,
        public readonly int $dataStartRow,
        public readonly int $dataEndRow,
        public readonly ?int $totalsRow,
        public readonly int $startCol,
        public readonly int $endCol,
        public readonly string $dataRange,
        public readonly string $tableRange,
    ) {}
}
