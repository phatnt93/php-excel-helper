<?php

declare(strict_types=1);

namespace ExcelHelper;

use ExcelHelper\Config\WriteConfigParser;
use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\I18n\Translator;
use ExcelHelper\Reader\ExcelReader;
use ExcelHelper\Reader\ReaderBuilder;
use ExcelHelper\Reader\ReadResult;
use ExcelHelper\Reader\ReadSummary;
use ExcelHelper\Reader\SheetInfo;
use ExcelHelper\Types\TypeRegistry;
use ExcelHelper\Writer\ExcelOutput;
use ExcelHelper\Writer\ExcelWriter;
use ExcelHelper\Writer\WriterBuilder;

/**
 * Static facade for php-excel-helper.
 */
final class Excel
{
    /**
     * Set the active translation locale ('en', 'vi').
     */
    public static function setLocale(string $locale): void
    {
        Translator::setLocale($locale);
    }

    /**
     * Get the currently active locale.
     */
    public static function getLocale(): string
    {
        return Translator::getLocale();
    }

    /**
     * Register a custom column type.
     */
    public static function registerType(TypeInterface $type, bool $replace = false): void
    {
        TypeRegistry::register($type, $replace);
    }
    /**
     * Reset global defaults (locale, registries) to initial clean state.
     * Useful between test cases.
     */
    public static function resetDefaults(): void
    {
        Translator::reset();
        TypeRegistry::reset();
    }

    /**
     * Read an Excel or CSV file using an array configuration.
     *
     * @param array<string, mixed> $options
     */
    public static function read(string $path, array $options = []): ReadResult
    {
        return ExcelReader::read($path, $options);
    }
    /**
     * Stream rows from an Excel (XLSX) or CSV file one-by-one as a Generator.
     * Provides low-memory row streaming (O(1) memory for UTF-8 CSV).
     *
     * @param array<string, mixed> $options
     * @return \Generator<int, array<string, mixed>>
     */
    public static function stream(string $path, array $options = []): \Generator
    {
        $opts = \ExcelHelper\Config\ReadConfigParser::parse($options);
        $inspection = \ExcelHelper\Reader\FileInspector::inspect($path, $opts);
        $format = $inspection['format'];

        if ($format === 'csv') {
            return \ExcelHelper\Reader\CsvStreamReader::stream($path, $opts);
        }

        if ($format === 'xlsx') {
            return \ExcelHelper\Reader\FastXlsxStreamReader::stream($path, $opts);
        }

        throw \ExcelHelper\Exceptions\UnsupportedFormatException::withMessage(sprintf(
            'Streaming is only supported for "xlsx" and "csv" formats. Provided: "%s". Use Excel::read() or Excel::chunk() for "%s".',
            $format,
            $format,
        ));
    }
    /**
     * Begin a fluent reader builder.
     */
    public static function reader(string $path): ReaderBuilder
    {
        return new ReaderBuilder($path);
    }

    /**
     * Read an Excel or CSV file in chunks.
     *
     * @param array<string, mixed>|callable $options
     * @param int|array<string, mixed>|callable $size
     */
    public static function chunk(
        string $path,
        array|callable $options,
        int|array|callable $size = 1000,
        ?callable $callback = null,
    ): ReadSummary {
        if (is_callable($options)) {
            $cb = $options;
            /** @var array<string, mixed> $opts */
            $opts = is_array($size) ? $size : [];
            $chunkSize = 1000;

            return ExcelReader::chunk($path, $opts, $chunkSize, $cb);
        }

        $opts = $options;
        $chunkSize = is_int($size) ? $size : 1000;
        /** @var callable(list<array<string, mixed>>, \ExcelHelper\Reader\ChunkContext): (bool|void) $cb */
        $cb = $callback ?? (is_callable($size) ? $size : static fn() => true);

        return ExcelReader::chunk($path, $opts, $chunkSize, $cb);
    }

    /**
     * List all worksheet names in a spreadsheet file without loading the entire document.
     *
     * @return list<string>
     */
    public static function sheetNames(string $path): array
    {
        return ExcelReader::sheetNames($path);
    }

    /**
     * Inspect sheet information (names, rows, columns) without loading entire sheets.
     *
     * @return list<SheetInfo>
     */
    public static function info(string $path): array
    {
        return ExcelReader::info($path);
    }

    /**
     * Extract header row from a file.
     *
     * @return array<string, string>
     */
    public static function headers(string $path, int $row = 1, int|string $sheet = 0): array
    {
        return ExcelReader::headers($path, $row, $sheet);
    }
    /**
     * Open an Excel spreadsheet template to fill placeholders, cell values, and tables.
     */
    public static function template(string $templatePath): \ExcelHelper\Template\TemplateProcessor
    {
        return new \ExcelHelper\Template\TemplateProcessor($templatePath);
    }

    /**
     * Build an export workbook from an array configuration.
     *
     * @param array<string, mixed> $config
     * @return ExcelOutput
     */
    public static function write(array $config = []): ExcelOutput
    {
        $def = WriteConfigParser::parse($config);
        $spreadsheet = ExcelWriter::build($def);

        return new ExcelOutput($spreadsheet, $def);
    }

    /**
     * Begin a fluent writer builder.
     */
    public static function writer(): WriterBuilder
    {
        return new WriterBuilder();
    }
}
