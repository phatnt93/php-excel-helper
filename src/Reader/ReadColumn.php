<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Immutable value object defining a column to be read and mapped.
 */
final class ReadColumn
{
    /**
     * @param string|int|null $column Column letter ('A'), 1-based index (1), or null when mapped by header
     * @param string|list<string>|null $header Header name(s) to match
     * @param string $key Key in the returned row array
     * @param string $type Registered type identifier
     * @param ?string $format Output format
     * @param string|list<string>|null $inputFormat Input format(s) for parsing
     * @param mixed $default Default value when empty or on error
     * @param bool $required True if an empty cell should trigger a required error
     * @param ?bool $trim Column-level trim override
     * @param ?array<string, mixed> $map Value mapping dictionary
     * @param bool $mapStrict True if unmapped value should trigger not_in_map error
     * @param (callable(mixed, RowContext): mixed)|null $transform Transform callback
     * @param (callable(mixed, RowContext): (bool|string))|null $validate Validation callback
     * @param list<string>|null $trueValues
     * @param list<string>|null $falseValues
     * @param bool $optional True if header column may be omitted from file without error
     * @param array<string, mixed> $options Additional options passed to TypeInterface
     */
    public function __construct(
        public readonly string $key,
        public readonly string|int|null $column = null,
        public readonly string|array|null $header = null,
        public readonly string $type = 'string',
        public readonly ?string $format = null,
        public readonly string|array|null $inputFormat = null,
        public readonly mixed $default = null,
        public readonly bool $required = false,
        public readonly ?bool $trim = null,
        public readonly ?array $map = null,
        public readonly bool $mapStrict = false,
        public readonly mixed $transform = null,
        public readonly mixed $validate = null,
        public readonly ?array $trueValues = null,
        public readonly ?array $falseValues = null,
        public readonly bool $optional = false,
        public readonly array $options = [],
    ) {}

    /**
     * Create a clone with an assigned resolved column index.
     */
    public function withResolvedColumn(?int $columnIndex): self
    {
        return new self(
            key: $this->key,
            column: $columnIndex,
            header: $this->header,
            type: $this->type,
            format: $this->format,
            inputFormat: $this->inputFormat,
            default: $this->default,
            required: $this->required,
            trim: $this->trim,
            map: $this->map,
            mapStrict: $this->mapStrict,
            transform: $this->transform,
            validate: $this->validate,
            trueValues: $this->trueValues,
            falseValues: $this->falseValues,
            optional: $this->optional,
            options: $this->options,
        );
    }
}
