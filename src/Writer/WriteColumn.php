<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

/**
 * Immutable value object defining a column to be exported.
 */
final class WriteColumn
{
    /**
     * @param string $key Key in the data row
     * @param ?string $header Header display text
     * @param string $type Registered type identifier ('auto', 'string', 'number', 'date', etc.)
     * @param ?string $format Number format code or date format string
     * @param int|float|string|null $width Column width in characters or 'auto'
     * @param array<string, mixed>|null $style Column-specific style configuration
     * @param (callable(mixed, array<string, mixed>): mixed)|null $value Value accessor callback
     * @param mixed $default Fallback value when null
     * @param array<string, mixed>|null $labels Boolean label mapping [true => 'Yes', false => 'No']
     * @param array<string, mixed>|null $validation Data validation configuration
     * @param list<array<string, mixed>>|array<string, mixed>|null $conditional Conditional formatting rules
     * @param array<string, mixed> $options Additional options
     */
    public function __construct(
        public readonly string $key,
        public readonly ?string $header = null,
        public readonly string $type = 'auto',
        public readonly ?string $format = null,
        public readonly int|float|string|null $width = null,
        public readonly ?array $style = null,
        public readonly mixed $value = null,
        public readonly mixed $default = null,
        public readonly ?array $labels = null,
        public readonly ?string $align = null,
        public readonly ?bool $wrap = null,
        public readonly ?array $validation = null,
        public readonly ?array $conditional = null,
        public readonly array $options = [],
    ) {}
}
