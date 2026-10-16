<?php

declare(strict_types=1);

namespace ExcelHelper\Config;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Reader\ReadColumn;
use ExcelHelper\Reader\ReadOptions;
use ExcelHelper\Support\Coordinate;
use ExcelHelper\Support\SizeParser;
use ExcelHelper\Types\TypeRegistry;

/**
 * Validates raw reader array configuration and produces an immutable ReadOptions value object.
 */
final class ReadConfigParser
{
    private const ALLOWED_ROOT_KEYS = [
        'format',
        'allowed_formats',
        'sheet',
        'header_row',
        'start_row',
        'end_row',
        'start_column',
        'columns',
        'skip_empty_rows',
        'stop_at_empty_row',
        'trim',
        'on_error',
        'row_number_key',
        'calculate_formulas',
        'max_rows',
        'max_file_size',
        'date_format',
        'datetime_format',
        'csv',
        'locale',
    ];

    private const ALLOWED_COLUMN_KEYS = [
        'key',
        'column',
        'header',
        'type',
        'format',
        'input_format',
        'default',
        'required',
        'trim',
        'map',
        'map_strict',
        'transform',
        'validate',
        'true_values',
        'false_values',
        'optional',
    ];

    private const VALID_ON_ERROR = ['collect', 'skip_row', 'throw'];

    /**
     * @param array<string, mixed> $config
     *
     * @throws InvalidConfigException
     */
    public static function parse(array $config): ReadOptions
    {
        $rootPath = ConfigPath::root('options');
        ConfigPath::assertKnownKeys($config, self::ALLOWED_ROOT_KEYS, $rootPath);

        $format = (string) ($config['format'] ?? 'auto');

        $rawAllowed = $config['allowed_formats'] ?? ReadOptions::DEFAULT_ALLOWED_FORMATS;
        if (!is_array($rawAllowed)) {
            throw InvalidConfigException::atPath($rootPath->append('allowed_formats'), 'Must be a list of format strings.');
        }
        $allowedFormats = array_values(array_map('strval', $rawAllowed));

        $sheet = $config['sheet'] ?? 0;
        if (!is_int($sheet) && !is_string($sheet)) {
            throw InvalidConfigException::atPath($rootPath->append('sheet'), 'Sheet must be an integer index (0-based) or string name.');
        }

        /** @var ?int $headerRow */
        $headerRow = isset($config['header_row']) ? (int) $config['header_row'] : null;
        if ($headerRow !== null && $headerRow < 1) {
            throw InvalidConfigException::atPath($rootPath->append('header_row'), 'Header row must be >= 1.');
        }

        $startRow = isset($config['start_row'])
            ? (int) $config['start_row']
            : ($headerRow !== null ? $headerRow + 1 : 1);

        if ($startRow < 1) {
            throw InvalidConfigException::atPath($rootPath->append('start_row'), 'Start row must be >= 1.');
        }

        if ($headerRow !== null && $headerRow >= $startRow) {
            throw InvalidConfigException::atPath(
                $rootPath->append('start_row'),
                sprintf('Start row (%d) must be greater than header row (%d).', $startRow, $headerRow),
            );
        }

        /** @var ?int $endRow */
        $endRow = isset($config['end_row']) ? (int) $config['end_row'] : null;
        if ($endRow !== null && $endRow < $startRow) {
            throw InvalidConfigException::atPath(
                $rootPath->append('end_row'),
                sprintf('End row (%d) cannot be less than start row (%d).', $endRow, $startRow),
            );
        }

        $startColumn = $config['start_column'] ?? 'A';
        if (!is_string($startColumn) && !is_int($startColumn)) {
            throw InvalidConfigException::atPath($rootPath->append('start_column'), 'Start column must be a column letter or integer.');
        }

        $onError = (string) ($config['on_error'] ?? 'collect');
        if (!in_array($onError, self::VALID_ON_ERROR, true)) {
            throw InvalidConfigException::atPath(
                $rootPath->append('on_error'),
                sprintf('Invalid on_error "%s". Allowed values: %s.', $onError, implode(', ', self::VALID_ON_ERROR)),
            );
        }

        $maxFileSize = SizeParser::parse($config['max_file_size'] ?? null);

        /** @var ?int $maxRows */
        $maxRows = isset($config['max_rows']) ? (int) $config['max_rows'] : null;
        if ($maxRows !== null && $maxRows < 1) {
            throw InvalidConfigException::atPath($rootPath->append('max_rows'), 'Max rows must be >= 1.');
        }

        $rawColumns = $config['columns'] ?? [];
        if (!is_array($rawColumns)) {
            throw InvalidConfigException::atPath($rootPath->append('columns'), 'Columns must be an array.');
        }

        $columns = self::parseColumns($rawColumns, $startColumn, $headerRow, $rootPath->append('columns'));

        $rawCsv = $config['csv'] ?? [];
        /** @var array<string, mixed> $csv */
        $csv = is_array($rawCsv) ? $rawCsv : [];

        return new ReadOptions(
            format: $format,
            allowedFormats: $allowedFormats,
            sheet: $sheet,
            headerRow: $headerRow,
            startRow: $startRow,
            endRow: $endRow,
            startColumn: $startColumn,
            columns: $columns,
            skipEmptyRows: (bool) ($config['skip_empty_rows'] ?? true),
            stopAtEmptyRow: (bool) ($config['stop_at_empty_row'] ?? false),
            trim: (bool) ($config['trim'] ?? true),
            onError: $onError,
            rowNumberKey: isset($config['row_number_key']) ? (string) $config['row_number_key'] : null,
            calculateFormulas: (bool) ($config['calculate_formulas'] ?? false),
            maxRows: $maxRows,
            maxFileSize: $maxFileSize,
            dateFormat: array_key_exists('date_format', $config) ? ($config['date_format'] !== null ? (string) $config['date_format'] : null) : 'Y-m-d',
            datetimeFormat: array_key_exists('datetime_format', $config) ? ($config['datetime_format'] !== null ? (string) $config['datetime_format'] : null) : 'Y-m-d H:i:s',
            csv: $csv,
            locale: isset($config['locale']) ? (string) $config['locale'] : null,
        );
    }

    /**
     * @param array<array-key, mixed> $rawColumns
     * @return list<ReadColumn>
     */
    private static function parseColumns(
        array $rawColumns,
        string|int $startColumn,
        ?int $headerRow,
        ConfigPath $path,
    ): array {
        if ($rawColumns === []) {
            return [];
        }

        $isList = array_is_list($rawColumns);
        $hasHeaders = false;
        $hasPositions = false;

        if ($isList) {
            foreach ($rawColumns as $idx => $item) {
                if ($item === null) {
                    continue;
                }
                if (is_array($item) && isset($item['header'])) {
                    $hasHeaders = true;
                } else {
                    $hasPositions = true;
                }
            }

            if ($hasHeaders && $hasPositions) {
                throw InvalidConfigException::atPath(
                    $path,
                    'Cannot mix sequential column definitions with header-based column definitions in the same list.',
                );
            }
        }

        $columns = [];
        $seenKeys = [];
        $seenPositions = [];

        if ($isList) {
            $baseColIndex = Coordinate::columnIndex($startColumn);
            $offset = 0;

            foreach ($rawColumns as $idx => $rawDef) {
                if ($rawDef === null) {
                    $offset++;
                    continue;
                }

                $subPath = $path->append($idx);
                $colDef = self::normalizeColumnDef($rawDef, $subPath);

                if ($hasHeaders) {
                    if ($headerRow === null) {
                        throw InvalidConfigException::atPath(
                            $subPath,
                            'Header-based column definitions require "header_row" to be configured.',
                        );
                    }
                    $readCol = self::buildColumn($colDef, null, $subPath);
                } else {
                    $assignedCol = $baseColIndex + $offset;
                    if (isset($seenPositions[$assignedCol])) {
                        throw InvalidConfigException::atPath($subPath, sprintf('Duplicate column index %d.', $assignedCol));
                    }
                    $seenPositions[$assignedCol] = true;
                    $readCol = self::buildColumn($colDef, $assignedCol, $subPath);
                    $offset++;
                }

                if (isset($seenKeys[$readCol->key])) {
                    throw InvalidConfigException::atPath($subPath, sprintf('Duplicate column key "%s".', $readCol->key));
                }
                $seenKeys[$readCol->key] = true;

                $columns[] = $readCol;
            }
        } else {
            // Associative array: keys are column letters or indices
            foreach ($rawColumns as $colKey => $rawDef) {
                $subPath = $path->append((string) $colKey);
                $colIndex = Coordinate::columnIndex($colKey);

                if (isset($seenPositions[$colIndex])) {
                    throw InvalidConfigException::atPath($subPath, sprintf('Duplicate column "%s".', (string) $colKey));
                }
                $seenPositions[$colIndex] = true;

                $colDef = self::normalizeColumnDef($rawDef, $subPath);

                if (isset($colDef['header'])) {
                    throw InvalidConfigException::atPath(
                        $subPath,
                        'Cannot specify "header" when columns are keyed by column letters or numbers.',
                    );
                }

                $readCol = self::buildColumn($colDef, $colIndex, $subPath);

                if (isset($seenKeys[$readCol->key])) {
                    throw InvalidConfigException::atPath($subPath, sprintf('Duplicate column key "%s".', $readCol->key));
                }
                $seenKeys[$readCol->key] = true;

                $columns[] = $readCol;
            }
        }

        return $columns;
    }

    /**
     * @return array<string, mixed>
     */
    private static function normalizeColumnDef(mixed $rawDef, ConfigPath $path): array
    {
        if (is_string($rawDef)) {
            $trimmed = trim($rawDef);
            if ($trimmed === '') {
                throw InvalidConfigException::atPath($path, 'Column key cannot be an empty string.');
            }

            return ['key' => $trimmed, 'type' => 'string'];
        }

        if (is_array($rawDef)) {
            return $rawDef;
        }

        throw InvalidConfigException::atPath($path, 'Column definition must be a string key or an array.');
    }

    /**
     * @param array<string, mixed> $colDef
     */
    private static function buildColumn(array $colDef, ?int $column, ConfigPath $path): ReadColumn
    {
        ConfigPath::assertKnownKeys($colDef, self::ALLOWED_COLUMN_KEYS, $path);

        $key = $colDef['key'] ?? null;
        if (!is_string($key) || trim($key) === '') {
            throw InvalidConfigException::atPath($path->append('key'), 'Column "key" is required and must be a non-empty string.');
        }

        $type = (string) ($colDef['type'] ?? 'string');
        if (!TypeRegistry::has($type)) {
            throw InvalidConfigException::atPath(
                $path->append('type'),
                sprintf('Unknown column type "%s".', $type),
            );
        }

        $header = $colDef['header'] ?? null;
        if ($header !== null && !is_string($header) && !is_array($header)) {
            throw InvalidConfigException::atPath($path->append('header'), '"header" must be a string or list of strings.');
        }

        $transform = $colDef['transform'] ?? null;
        if ($transform !== null && !is_callable($transform)) {
            throw InvalidConfigException::atPath($path->append('transform'), '"transform" must be callable.');
        }

        $validate = $colDef['validate'] ?? null;
        if ($validate !== null && !is_callable($validate)) {
            throw InvalidConfigException::atPath($path->append('validate'), '"validate" must be callable.');
        }

        $map = $colDef['map'] ?? null;
        if ($map !== null && !is_array($map)) {
            throw InvalidConfigException::atPath($path->append('map'), '"map" must be an array.');
        }
        $headerVal = null;
        if (isset($colDef['header'])) {
            if (is_array($colDef['header'])) {
                $headerVal = array_values(array_map('strval', $colDef['header']));
            } elseif (is_string($colDef['header'])) {
                $headerVal = $colDef['header'];
            }
        }

        $trueValues = isset($colDef['true_values']) && is_array($colDef['true_values'])
            ? array_values(array_map('strval', $colDef['true_values']))
            : null;

        $falseValues = isset($colDef['false_values']) && is_array($colDef['false_values'])
            ? array_values(array_map('strval', $colDef['false_values']))
            : null;

        return new ReadColumn(
            key: trim($key),
            column: $column,
            header: $headerVal,
            type: $type,
            format: isset($colDef['format']) ? (string) $colDef['format'] : null,
            inputFormat: $colDef['input_format'] ?? null,
            default: $colDef['default'] ?? null,
            required: (bool) ($colDef['required'] ?? false),
            trim: isset($colDef['trim']) ? (bool) $colDef['trim'] : null,
            map: $map,
            mapStrict: (bool) ($colDef['map_strict'] ?? false),
            transform: $transform,
            validate: $validate,
            trueValues: $trueValues,
            falseValues: $falseValues,
            optional: (bool) ($colDef['optional'] ?? false),
            options: $colDef,
        );
    }
}
