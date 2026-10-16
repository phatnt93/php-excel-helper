<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Metadata about a single worksheet without loading cell data into memory.
 */
final class SheetInfo
{
    public function __construct(
        public readonly string $name,
        public readonly int $totalRows,
        public readonly int $totalColumns,
        public readonly string $lastColumnLetter,
        public readonly int $lastColumnIndex,
    ) {}

    /**
     * @return array{
     *     name: string,
     *     totalRows: int,
     *     totalColumns: int,
     *     lastColumnLetter: string,
     *     lastColumnIndex: int
     * }
     */
    public function toArray(): array
    {
        return [
            'name' => $this->name,
            'totalRows' => $this->totalRows,
            'totalColumns' => $this->totalColumns,
            'lastColumnLetter' => $this->lastColumnLetter,
            'lastColumnIndex' => $this->lastColumnIndex,
        ];
    }
}
