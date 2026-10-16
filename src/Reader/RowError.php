<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Immutable value object representing a row-level data error.
 */
final class RowError
{
    public function __construct(
        public readonly int $row,
        public readonly string $column,
        public readonly string $key,
        public readonly string $code,
        public readonly mixed $value,
        public readonly string $message,
    ) {}

    /**
     * @return array{
     *     row: int,
     *     column: string,
     *     key: string,
     *     code: string,
     *     value: mixed,
     *     message: string
     * }
     */
    public function toArray(): array
    {
        return [
            'row' => $this->row,
            'column' => $this->column,
            'key' => $this->key,
            'code' => $this->code,
            'value' => $this->value,
            'message' => $this->message,
        ];
    }
}
