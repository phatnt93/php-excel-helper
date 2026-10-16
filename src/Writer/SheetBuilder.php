<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

/**
 * Fluent builder for configuring a single worksheet.
 */
final class SheetBuilder
{
    /**
     * @var array<string, mixed>
     */
    private array $config = [];

    public function __construct(string $name = 'Sheet1')
    {
        $this->config['name'] = $name;
    }

    public function name(string $name): self
    {
        $this->config['name'] = $name;

        return $this;
    }

    /**
     * @param array<string, mixed>|null $style
     */
    public function title(string $title, ?array $style = null): self
    {
        $this->config['title'] = $title;
        if ($style !== null) {
            $this->config['title_style'] = $style;
        }

        return $this;
    }

    public function startCell(string $cell): self
    {
        $this->config['start_cell'] = $cell;

        return $this;
    }

    /**
     * Define a column to export.
     *
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
        $columns = $this->config['columns'] ?? [];

        $colDef = array_merge([
            'key' => $key,
            'header' => $header ?? $key,
            'type' => $type,
            'format' => $format,
            'width' => $width,
            'style' => $style,
            'value' => $value,
            'default' => $default,
            'labels' => $labels,
        ], $options);

        $columns[] = $colDef;
        $this->config['columns'] = $columns;

        return $this;
    }

    /**
     * @param array<array-key, mixed> $columns
     */
    public function columns(array $columns): self
    {
        $this->config['columns'] = $columns;

        return $this;
    }

    /**
     * @param list<string> $headers
     */
    public function headers(array $headers): self
    {
        $this->config['headers'] = $headers;

        return $this;
    }

    /**
     * @param iterable<mixed> $data
     */
    public function data(iterable $data): self
    {
        $this->config['data'] = $data;

        return $this;
    }

    /**
     * @param array<string, string> $totals
     */
    public function totals(array $totals, ?string $label = null): self
    {
        $this->config['totals'] = $totals;
        if ($label !== null) {
            $this->config['totals_label'] = $label;
        }

        return $this;
    }

    public function showHeader(bool $show = true): self
    {
        $this->config['show_header'] = $show;

        return $this;
    }
    /**
     * @param array<string, mixed> $style
     */
    public function headerStyle(array $style): self
    {
        $this->config['header_style'] = $style;

        return $this;
    }

    /**
     * @param array<string, mixed> $style
     */
    public function rowStyle(array $style): self
    {
        $this->config['row_style'] = $style;

        return $this;
    }

    /**
     * @param bool|array<string, mixed> $zebra
     */
    public function zebra(bool|array $zebra = true): self
    {
        $this->config['zebra'] = $zebra;

        return $this;
    }

    /**
     * @param bool|string|array<string, mixed> $borders
     */
    public function borders(bool|string|array $borders = true): self
    {
        $this->config['borders'] = $borders;

        return $this;
    }

    public function freezeHeader(bool $freeze = true): self
    {
        $this->config['freeze_header'] = $freeze;

        return $this;
    }

    public function autoFilter(bool $filter = true): self
    {
        $this->config['auto_filter'] = $filter;

        return $this;
    }

    public function autoSize(?bool $autoSize = true, int $maxRows = 5000): self
    {
        $this->config['auto_size'] = $autoSize;
        $this->config['auto_size_max_rows'] = $maxRows;

        return $this;
    }

    public function rowHeight(?float $height): self
    {
        $this->config['row_height'] = $height;

        return $this;
    }

    public function merge(string $range): self
    {
        $merges = $this->config['merges'] ?? [];
        $merges[] = $range;
        $this->config['merges'] = $merges;

        return $this;
    }

    /**
     * @param list<string> $ranges
     */
    public function merges(array $ranges): self
    {
        $existing = $this->config['merges'] ?? [];
        $this->config['merges'] = array_merge($existing, $ranges);

        return $this;
    }
    /**
     * @param array<string, mixed>|ImageDefinition|string $image
     */
    public function image(
        string|array|ImageDefinition $image,
        string $cell = 'A1',
        ?int $width = null,
        ?int $height = null,
    ): self {
        $images = $this->config['images'] ?? [];
        if ($image instanceof ImageDefinition) {
            $images[] = [
                'path' => $image->path,
                'cell' => $image->cell,
                'width' => $image->width,
                'height' => $image->height,
                'offset_x' => $image->offsetX,
                'offset_y' => $image->offsetY,
                'name' => $image->name,
                'description' => $image->description,
            ];
        } elseif (is_array($image)) {
            $images[] = $image;
        } else {
            $images[] = [
                'path' => $image,
                'cell' => $cell,
                'width' => $width,
                'height' => $height,
            ];
        }
        $this->config['images'] = $images;

        return $this;
    }

    /**
     * @param list<array<string, mixed>|ImageDefinition|string> $images
     */
    public function images(array $images): self
    {
        $existing = $this->config['images'] ?? [];
        $this->config['images'] = array_merge($existing, $images);

        return $this;
    }

    /**
     * @param callable(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, SheetLayout): void $tap
     */
    public function tap(callable $tap): self
    {
        $this->config['tap'] = $tap;

        return $this;
    }

    /**
     * @return array<string, mixed>
     */
    public function toConfig(): array
    {
        return $this->config;
    }
}
