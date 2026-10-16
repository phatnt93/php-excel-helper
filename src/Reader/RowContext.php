<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Contextual information provided to transform and validate callbacks during row mapping.
 */
final class RowContext
{
    /**
     * @param int $row 1-based Excel row number
     * @param array<int, mixed> $rawRow Raw cell values indexed by 1-based column index
     * @param array<string, mixed> $mappedRow Mapped values processed so far
     */
    public function __construct(
        public readonly int $row,
        public readonly array $rawRow,
        public readonly array $mappedRow,
    ) {}
}
