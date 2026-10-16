<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

/**
 * Immutable value object holding parsed reading options.
 */
final class ReadOptions
{
    /**
     * @var list<string>
     */
    public const DEFAULT_ALLOWED_FORMATS = ['xlsx', 'xls', 'ods', 'csv'];

    /**
     * @param string $format 'auto'|'xlsx'|'xls'|'ods'|'csv'
     * @param list<string> $allowedFormats Allowed format whitelist
     * @param int|string $sheet 0-based sheet index or sheet name
     * @param ?int $headerRow 1-based header row number
     * @param int $startRow 1-based first data row number
     * @param ?int $endRow 1-based last data row number (inclusive)
     * @param string|int $startColumn Starting column letter or index for sequential list
     * @param list<ReadColumn> $columns Column definitions
     * @param bool $skipEmptyRows Skip rows where all mapped cells are empty
     * @param bool $stopAtEmptyRow Stop reading upon hitting the first empty row
     * @param bool $trim Trim string values by default
     * @param string $onError 'collect'|'skip_row'|'throw'
     * @param ?string $rowNumberKey Key to store 1-based row number in output row
     * @param bool $calculateFormulas If false, returns cached formula values
     * @param ?int $maxRows Maximum allowed rows
     * @param ?int $maxFileSize Maximum allowed file size in bytes
     * @param ?string $dateFormat Default date output format
     * @param ?string $datetimeFormat Default datetime output format
     * @param array<string, mixed> $csv CSV parser options
     * @param ?string $locale Error translation locale
     */
    public function __construct(
        public readonly string $format = 'auto',
        public readonly array $allowedFormats = self::DEFAULT_ALLOWED_FORMATS,
        public readonly int|string $sheet = 0,
        public readonly ?int $headerRow = null,
        public readonly int $startRow = 1,
        public readonly ?int $endRow = null,
        public readonly string|int $startColumn = 'A',
        public readonly array $columns = [],
        public readonly bool $skipEmptyRows = true,
        public readonly bool $stopAtEmptyRow = false,
        public readonly bool $trim = true,
        public readonly string $onError = 'collect',
        public readonly ?string $rowNumberKey = null,
        public readonly bool $calculateFormulas = false,
        public readonly ?int $maxRows = null,
        public readonly ?int $maxFileSize = null,
        public readonly ?string $dateFormat = 'Y-m-d',
        public readonly ?string $datetimeFormat = 'Y-m-d H:i:s',
        public readonly array $csv = [],
        public readonly ?string $locale = null,
    ) {}
}
