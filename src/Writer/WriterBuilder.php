<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

use ExcelHelper\Excel;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Fluent builder for configuring and executing spreadsheet exports.
 */
final class WriterBuilder
{
    private string $format = 'xlsx';

    /**
     * @var array<string, string>
     */
    private array $properties = [];

    /**
     * @var array<string, mixed>
     */
    private array $defaults = [];

    /**
     * @var array<string, mixed>
     */
    private array $csv = [];

    /**
     * @var list<array<string, mixed>>
     */
    private array $sheets = [];
    private ?string $locale = null;

    private SheetBuilder $singleSheet;

    public function __construct()
    {
        $this->singleSheet = new SheetBuilder('Sheet1');
    }

    public function format(string $format): self
    {
        $this->format = $format;

        return $this;
    }

    public function property(string $name, string $value): self
    {
        $this->properties[$name] = $value;

        return $this;
    }

    /**
     * @param array<string, string> $properties
     */
    public function properties(array $properties): self
    {
        $this->properties = array_merge($this->properties, $properties);

        return $this;
    }

    public function default(string $name, mixed $value): self
    {
        $this->defaults[$name] = $value;

        return $this;
    }

    /**
     * @param array<string, mixed> $defaults
     */
    public function defaults(array $defaults): self
    {
        $this->defaults = array_merge($this->defaults, $defaults);

        return $this;
    }
    public function locale(?string $locale): self
    {
        $this->locale = $locale;

        return $this;
    }

    public function showHeader(bool $show = true): self
    {
        $this->singleSheet->showHeader($show);

        return $this;
    }

    /**
     * @param array<string, mixed> $csvOptions
     */
    public function csv(array $csvOptions): self
    {
        $this->csv = $csvOptions;

        return $this;
    }

    /**
     * Configure a sheet using a closure callback.
     *
     * @param (callable(SheetBuilder): mixed)|null $closure
     */
    public function sheet(string $name, ?callable $closure = null): self
    {
        $builder = new SheetBuilder($name);

        if ($closure !== null) {
            $closure($builder);
            $this->sheets[] = $builder->toConfig();
        } else {
            $this->singleSheet = $builder;
        }

        return $this;
    }

    // Single-sheet convenience methods

    /**
     * @param array<string, mixed>|null $style
     */
    public function title(string $title, ?array $style = null): self
    {
        $this->singleSheet->title($title, $style);

        return $this;
    }

    public function startCell(string $cell): self
    {
        $this->singleSheet->startCell($cell);

        return $this;
    }

    /**
     * @param array<string, mixed>|null $style
     * @param array<string, mixed>|null $labels
     * @param array<string, mixed> $options
     */
    public function column(
        string $key,
        ?string $header = null,
        string $type = 'auto',
        ?string $format = null,
        int|float|string|null $width = null,
        ?array $style = null,
        ?callable $value = null,
        mixed $default = null,
        ?array $labels = null,
        array $options = [],
    ): self {
        $this->singleSheet->column(
            $key,
            $header,
            $type,
            $format,
            $width,
            $style,
            $value,
            $default,
            $labels,
            $options,
        );

        return $this;
    }

    /**
     * @param array<array-key, mixed> $columns
     */
    public function columns(array $columns): self
    {
        $this->singleSheet->columns($columns);

        return $this;
    }

    /**
     * @param list<string> $headers
     */
    public function headers(array $headers): self
    {
        $this->singleSheet->headers($headers);

        return $this;
    }

    /**
     * @param iterable<mixed> $data
     */
    public function data(iterable $data): self
    {
        $this->singleSheet->data($data);

        return $this;
    }

    /**
     * @param array<string, string> $totals
     */
    public function totals(array $totals, ?string $label = null): self
    {
        $this->singleSheet->totals($totals, $label);

        return $this;
    }

    /**
     * @param array<string, mixed> $style
     */
    public function headerStyle(array $style): self
    {
        $this->singleSheet->headerStyle($style);

        return $this;
    }

    /**
     * @param array<string, mixed> $style
     */
    public function rowStyle(array $style): self
    {
        $this->singleSheet->rowStyle($style);

        return $this;
    }

    /**
     * @param bool|array<string, mixed> $zebra
     */
    public function zebra(bool|array $zebra = true): self
    {
        $this->singleSheet->zebra($zebra);

        return $this;
    }

    /**
     * @param bool|string|array<string, mixed> $borders
     */
    public function borders(bool|string|array $borders = true): self
    {
        $this->singleSheet->borders($borders);

        return $this;
    }

    public function freezeHeader(bool $freeze = true): self
    {
        $this->singleSheet->freezeHeader($freeze);

        return $this;
    }

    public function autoFilter(bool $filter = true): self
    {
        $this->singleSheet->autoFilter($filter);

        return $this;
    }

    public function autoSize(?bool $autoSize = true, int $maxRows = 5000): self
    {
        $this->singleSheet->autoSize($autoSize, $maxRows);

        return $this;
    }

    public function rowHeight(?float $height): self
    {
        $this->singleSheet->rowHeight($height);

        return $this;
    }

    public function merge(string $range): self
    {
        $this->singleSheet->merge($range);

        return $this;
    }

    /**
     * @param list<string> $ranges
     */
    public function merges(array $ranges): self
    {
        $this->singleSheet->merges($ranges);

        return $this;
    }
    /**
     * @param array<string, mixed>|\ExcelHelper\Writer\ImageDefinition|string $image
     */
    public function image(
        string|array|\ExcelHelper\Writer\ImageDefinition $image,
        string $cell = 'A1',
        ?int $width = null,
        ?int $height = null,
    ): self {
        $this->singleSheet->image($image, $cell, $width, $height);

        return $this;
    }

    /**
     * @param list<array<string, mixed>|\ExcelHelper\Writer\ImageDefinition|string> $images
     */
    public function images(array $images): self
    {
        $this->singleSheet->images($images);

        return $this;
    }

    /**
     * @param callable(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, SheetLayout): void $tap
     */
    public function tap(callable $tap): self
    {
        $this->singleSheet->tap($tap);

        return $this;
    }

    /**
     * Build the export configuration array.
     *
     * @return array<string, mixed>
     */
    public function toConfig(): array
    {
        $base = [
            'format' => $this->format,
            'properties' => $this->properties,
            'defaults' => $this->defaults,
            'csv' => $this->csv,
        ];

        if ($this->locale !== null) {
            $base['locale'] = $this->locale;
        }

        if ($this->sheets !== []) {
            $base['sheets'] = $this->sheets;

            return $base;
        }

        return array_merge($this->singleSheet->toConfig(), $base);
    }

    /**
     * Build and return the ExcelOutput wrapper.
     */
    public function output(): ExcelOutput
    {
        return Excel::write($this->toConfig());
    }

    public function save(string $path): string
    {
        return $this->output()->save($path);
    }

    public function download(string $filename): void
    {
        $this->output()->download($filename);
    }

    public function toString(): string
    {
        return $this->output()->toString();
    }

    public function spreadsheet(): Spreadsheet
    {
        return $this->output()->spreadsheet();
    }
}
