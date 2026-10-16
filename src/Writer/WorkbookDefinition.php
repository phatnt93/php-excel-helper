<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

/**
 * Immutable value object holding configuration for writing an entire workbook.
 */
final class WorkbookDefinition
{
    /**
     * @var list<string>
     */
    public const SUPPORTED_FORMATS = ['xlsx', 'xls', 'ods', 'csv'];

    /**
     * @param string $format 'xlsx'|'xls'|'ods'|'csv'
     * @param array<string, string> $properties Document metadata properties
     * @param list<SheetDefinition> $sheets Worksheet definitions
     * @param array<string, mixed> $defaults Default options inherited by all sheets
     * @param array<string, mixed> $csv CSV writer options
     */
    public function __construct(
        public readonly string $format = 'xlsx',
        public readonly array $properties = [],
        public readonly array $sheets = [],
        public readonly array $defaults = [],
        public readonly array $csv = [],
        public readonly ?string $locale = null,
    ) {}
}
