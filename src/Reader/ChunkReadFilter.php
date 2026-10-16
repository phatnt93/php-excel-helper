<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

/**
 * Filter that limits reading to specific rows (and optional columns), reducing memory usage.
 */
final class ChunkReadFilter implements IReadFilter
{
    /**
     * @var array<string, bool>|null
     */
    private ?array $columns = null;

    /**
     * @param list<string>|null $allowedColumns List of column letters (e.g. ['A', 'B', 'E'])
     */
    public function __construct(
        private int $startRow = 1,
        private ?int $endRow = null,
        private readonly ?int $headerRow = null,
        ?array $allowedColumns = null,
    ) {
        if ($allowedColumns !== null && $allowedColumns !== []) {
            $this->columns = array_fill_keys(array_map('strtoupper', $allowedColumns), true);
        }
    }

    public function setRows(int $startRow, ?int $endRow): void
    {
        $this->startRow = $startRow;
        $this->endRow = $endRow;
    }

    public function readCell(string $columnAddress, int $row, string $worksheetName = ''): bool
    {
        if ($this->headerRow !== null && $row === $this->headerRow) {
            return true;
        }

        if ($row < $this->startRow) {
            return false;
        }

        if ($this->endRow !== null && $row > $this->endRow) {
            return false;
        }

        if ($this->columns !== null) {
            return isset($this->columns[strtoupper($columnAddress)]);
        }

        return true;
    }
}
