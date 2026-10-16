<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Result object returned by Excel::read().
 */
final class ReadResult
{
    /**
     * @var array<int, list<RowError>>|null
     */
    private ?array $errorsByRowCache = null;

    /**
     * @param list<array<string, mixed>> $rows
     * @param list<RowError> $errors
     * @param array<string, string> $headers
     */
    public function __construct(
        private readonly array $rows,
        private readonly array $errors = [],
        private readonly array $headers = [],
        private readonly string $sheetName = '',
    ) {}

    /**
     * @return list<array<string, mixed>>
     */
    public function rows(): array
    {
        return $this->rows;
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
     * @return array<int, list<RowError>>
     */
    public function errorsByRow(): array
    {
        if ($this->errorsByRowCache !== null) {
            return $this->errorsByRowCache;
        }

        $grouped = [];
        foreach ($this->errors as $error) {
            $grouped[$error->row][] = $error;
        }

        $this->errorsByRowCache = $grouped;

        return $grouped;
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

    public function rowCount(): int
    {
        return count($this->rows);
    }

    /**
     * @return array{
     *     rows: list<array<string, mixed>>,
     *     errors: list<array<string, mixed>>,
     *     headers: array<string, string>,
     *     sheetName: string,
     *     rowCount: int
     * }
     */
    public function toArray(): array
    {
        return [
            'rows' => $this->rows,
            'errors' => array_map(static fn(RowError $e): array => $e->toArray(), $this->errors),
            'headers' => $this->headers,
            'sheetName' => $this->sheetName,
            'rowCount' => $this->rowCount(),
        ];
    }
}
