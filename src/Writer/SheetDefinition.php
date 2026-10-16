<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

/**
 * Immutable value object holding configuration for writing a single worksheet.
 */
final class SheetDefinition
{
    /**
     * @param string $name Worksheet tab title
     * @param ?string $title Large banner title above the table
     * @param array<string, mixed>|null $titleStyle Style for the banner title
     * @param string $startCell Top-left cell where the sheet table begins (e.g. 'A1')
     * @param list<WriteColumn> $columns Column definitions
     * @param list<string>|null $headers Simple header strings (used if columns not specified)
     * @param iterable<mixed> $data Rows to export (array or Traversable/Generator)
     * @param array<string, string>|null $totals Column totals configuration (e.g. ['salary' => 'sum'])
     * @param ?string $totalsLabel Label for the totals row
     * @param array<string, mixed>|null $headerStyle Style configuration for the header row
     * @param array<string, mixed>|null $rowStyle Style configuration for data rows
     * @param bool|array<string, mixed> $zebra Alternating row background styling
     * @param bool|string|array<string, mixed> $borders Table borders styling
     * @param bool $freezeHeader Freeze panes below header
     * @param bool $autoFilter Enable auto filter on header row
     * @param ?bool $autoSize Auto size column widths (null = auto: true if rows <= autoSizeMaxRows)
     * @param int $autoSizeMaxRows Row count threshold to disable auto-sizing
     * @param ?float $rowHeight Custom row height
     * @param list<string> $merges Custom cell merges (e.g. ['A1:D1'])
     * @param list<ImageDefinition> $images Images or logos to insert into the worksheet
     * @param list<array<string, mixed>> $conditionalFormatting Conditional formatting rules
     * @param (callable(\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, SheetLayout): void)|null $tap Hook callback
     */
    public function __construct(
        public readonly string $name = 'Sheet1',
        public readonly ?string $title = null,
        public readonly ?array $titleStyle = null,
        public readonly string $startCell = 'A1',
        public readonly array $columns = [],
        public readonly ?array $headers = null,
        public readonly iterable $data = [],
        public readonly bool $showHeader = true,
        public readonly ?array $totals = null,
        public readonly ?string $totalsLabel = null,
        public readonly ?array $headerStyle = null,
        public readonly ?array $rowStyle = null,
        public readonly bool|array $zebra = false,
        public readonly bool|string|array $borders = false,
        public readonly bool $freezeHeader = true,
        public readonly bool $autoFilter = false,
        public readonly ?bool $autoSize = null,
        public readonly int $autoSizeMaxRows = 5000,
        public readonly ?float $rowHeight = null,
        public readonly array $merges = [],
        public readonly array $images = [],
        public readonly array $conditionalFormatting = [],
        public readonly mixed $tap = null,
    ) {}
}
