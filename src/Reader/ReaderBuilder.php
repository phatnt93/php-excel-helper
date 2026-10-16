<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use ExcelHelper\Support\Coordinate;

/**
 * Fluent builder for configuring and executing spreadsheet reads.
 */
final class ReaderBuilder
{
    /**
     * @var array<string, mixed>
     */
    private array $options = [];

    public function __construct(
        private readonly string $path,
    ) {}

    public function sheet(int|string $sheet): self
    {
        $this->options['sheet'] = $sheet;

        return $this;
    }

    public function headerRow(?int $row): self
    {
        $this->options['header_row'] = $row;

        return $this;
    }

    public function startRow(int $row): self
    {
        $this->options['start_row'] = $row;

        return $this;
    }

    public function endRow(?int $row): self
    {
        $this->options['end_row'] = $row;

        return $this;
    }

    public function startColumn(string|int $col): self
    {
        $this->options['start_column'] = $col;

        return $this;
    }

    /**
     * Define a column to read.
     *
     * Supports both:
     * - Keyed by letter: column('A', 'code') or column('A', 'birthday', type: 'date', format: 'd/m/Y')
     * - Config array: column('code', ['header' => 'Mã NV'])
     *
     * @param array<string, mixed>|string|null $configOrKey
     * @param array<string, mixed>|null $map
     * @param array<string, mixed> $options
     */
    public function column(
        string|int $colOrKey,
        string|array|null $configOrKey = null,
        string $type = 'string',
        ?string $format = null,
        mixed $default = null,
        bool $required = false,
        ?array $map = null,
        bool $mapStrict = false,
        ?callable $transform = null,
        ?callable $validate = null,
        bool $optional = false,
        array $options = [],
    ): self {
        $columns = $this->options['columns'] ?? [];

        if (is_array($configOrKey)) {
            $colConfig = array_merge(['key' => (string) $colOrKey], $configOrKey);
            $columns[] = $colConfig;
        } else {
            $key = $configOrKey !== null ? (string) $configOrKey : (string) $colOrKey;
            $colConfig = array_merge([
                'key' => $key,
                'type' => $type,
                'format' => $format,
                'default' => $default,
                'required' => $required,
                'map' => $map,
                'map_strict' => $mapStrict,
                'transform' => $transform,
                'validate' => $validate,
                'optional' => $optional,
            ], $options);

            // Check if colOrKey is a column letter or integer
            if (is_int($colOrKey) || Coordinate::isColumnLetter((string) $colOrKey)) {
                $columns[$colOrKey] = $colConfig;
            } else {
                $columns[] = $colConfig;
            }
        }

        $this->options['columns'] = $columns;

        return $this;
    }

    /**
     * @param array<array-key, mixed> $columns
     */
    public function columns(array $columns): self
    {
        $this->options['columns'] = $columns;

        return $this;
    }

    public function skipEmptyRows(bool $skip = true): self
    {
        $this->options['skip_empty_rows'] = $skip;

        return $this;
    }

    public function stopAtEmptyRow(bool $stop = true): self
    {
        $this->options['stop_at_empty_row'] = $stop;

        return $this;
    }

    public function trim(bool $trim = true): self
    {
        $this->options['trim'] = $trim;

        return $this;
    }

    public function onError(string $policy): self
    {
        $this->options['on_error'] = $policy;

        return $this;
    }

    public function rowNumberKey(?string $key): self
    {
        $this->options['row_number_key'] = $key;

        return $this;
    }

    public function calculateFormulas(bool $calc = true): self
    {
        $this->options['calculate_formulas'] = $calc;

        return $this;
    }

    public function maxRows(?int $rows): self
    {
        $this->options['max_rows'] = $rows;

        return $this;
    }

    public function maxFileSize(int|string|null $size): self
    {
        $this->options['max_file_size'] = $size;

        return $this;
    }

    public function dateFormat(?string $format): self
    {
        $this->options['date_format'] = $format;

        return $this;
    }

    public function datetimeFormat(?string $format): self
    {
        $this->options['datetime_format'] = $format;

        return $this;
    }

    /**
     * @param array<string, mixed> $csvOptions
     */
    public function csv(array $csvOptions): self
    {
        $this->options['csv'] = $csvOptions;

        return $this;
    }

    public function locale(?string $locale): self
    {
        $this->options['locale'] = $locale;

        return $this;
    }

    public function format(string $format): self
    {
        $this->options['format'] = $format;

        return $this;
    }

    /**
     * @param list<string> $formats
     */
    public function allowedFormats(array $formats): self
    {
        $this->options['allowed_formats'] = $formats;

        return $this;
    }

    /**
     * @return array<string, mixed>
     */
    public function toOptions(): array
    {
        return $this->options;
    }

    /**
     * Execute the read operation and return a ReadResult.
     */
    public function read(): ReadResult
    {
        return ExcelReader::read($this->path, $this->options);
    }

    /**
     * Execute reading in chunks.
     *
     * @param callable(list<array<string, mixed>>, ChunkContext): (bool|void) $callback
     */
    public function chunk(int $size, callable $callback): ReadSummary
    {
        return ExcelReader::chunk($this->path, $this->options, $size, $callback);
    }
}
