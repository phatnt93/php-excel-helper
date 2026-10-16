<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Context passed to chunk callbacks during batch reading.
 */
final class ChunkContext
{
    private bool $shouldStop = false;

    /**
     * @param list<RowError> $errors
     */
    public function __construct(
        public readonly int $chunkIndex,
        public readonly int $startRow,
        public readonly int $endRow,
        public readonly int $rowCount,
        public readonly int $totalRowsProcessed,
        private readonly array $errors = [],
    ) {}

    /**
     * @return list<RowError>
     */
    public function errors(): array
    {
        return $this->errors;
    }

    public function chunkIndex(): int
    {
        return $this->chunkIndex;
    }

    public function stop(): void
    {
        $this->shouldStop = true;
    }

    public function shouldStop(): bool
    {
        return $this->shouldStop;
    }
}
