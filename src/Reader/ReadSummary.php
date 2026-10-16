<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Summary statistics returned by Excel::chunk().
 */
final class ReadSummary
{
    /**
     * @param list<RowError> $errors
     * @param array<string, string> $headers
     */
    public function __construct(
        public readonly int $totalRows,
        public readonly int $totalChunks,
        public readonly array $errors = [],
        public readonly array $headers = [],
        public readonly string $sheetName = '',
    ) {}

    public function totalRows(): int
    {
        return $this->totalRows;
    }

    public function totalChunks(): int
    {
        return $this->totalChunks;
    }

    public function errorCount(): int
    {
        return count($this->errors);
    }

    /**
     * @return list<RowError>
     */
    public function errors(): array
    {
        return $this->errors;
    }

    public function hasErrors(): bool
    {
        return $this->errors !== [];
    }

    /**
     * @return array<string, string>
     */
    public function headers(): array
    {
        return $this->headers;
    }

    public function sheetName(): string
    {
        return $this->sheetName;
    }
}
