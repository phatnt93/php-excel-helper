<?php

declare(strict_types=1);

namespace ExcelHelper\Config;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Support\Coordinate;
use ExcelHelper\Types\TypeRegistry;
use ExcelHelper\Writer\ImageDefinition;
use ExcelHelper\Writer\SheetDefinition;
use ExcelHelper\Writer\WorkbookDefinition;
use ExcelHelper\Writer\WriteColumn;

/**
 * Validates raw writer array configuration and produces a WorkbookDefinition.
 */
final class WriteConfigParser
{
    private const WORKBOOK_ROOT_KEYS = [
        'format',
        'properties',
        'sheets',
        'defaults',
        'csv',
        'locale',
    ];
    private const SHEET_KEYS = [
        'name',
        'title',
        'title_style',
        'start_cell',
        'columns',
        'headers',
        'data',
        'totals',
        'totals_label',
        'show_header',
        'header_style',
        'row_style',
        'zebra',
        'borders',
        'freeze_header',
        'auto_filter',
        'auto_size',
        'auto_size_max_rows',
        'row_height',
        'merges',
        'images',
        'conditional_formatting',
        'tap',
    ];

    private const ALLOWED_COLUMN_KEYS = [
        'key',
        'header',
        'type',
        'format',
        'width',
        'align',
        'wrap',
        'style',
        'value',
        'default',
        'labels',
        'validation',
        'conditional',
    ];

    /**
     * @param array<string, mixed> $config
     *
     * @throws InvalidConfigException
     */
    public static function parse(array $config): WorkbookDefinition
    {
        $rootPath = ConfigPath::root('config');

        $format = (string) ($config['format'] ?? 'xlsx');
        $formatLower = strtolower(trim($format));

        if (!in_array($formatLower, WorkbookDefinition::SUPPORTED_FORMATS, true)) {
            throw InvalidConfigException::atPath(
                $rootPath->append('format'),
                sprintf('Unsupported export format "%s". Supported formats are: %s.', $format, implode(', ', WorkbookDefinition::SUPPORTED_FORMATS)),
            );
        }

        /** @var array<string, string> $properties */
        $properties = isset($config['properties']) && is_array($config['properties']) ? $config['properties'] : [];
        /** @var array<string, mixed> $defaults */
        $defaults = isset($config['defaults']) && is_array($config['defaults']) ? $config['defaults'] : [];
        /** @var array<string, mixed> $csv */
        $csv = isset($config['csv']) && is_array($config['csv']) ? $config['csv'] : [];
        $locale = isset($config['locale']) && is_string($config['locale']) ? trim($config['locale']) : null;
        $sheets = [];

        if (isset($config['sheets'])) {
            // Multi-sheet mode
            if (!is_array($config['sheets'])) {
                throw InvalidConfigException::atPath($rootPath->append('sheets'), '"sheets" must be a list of sheet configurations.');
            }

            // Ensure no root sheet keys were mixed with sheets
            foreach (self::SHEET_KEYS as $sheetKey) {
                if (array_key_exists($sheetKey, $config)) {
                    throw InvalidConfigException::atPath(
                        $rootPath->append($sheetKey),
                        sprintf('Cannot specify root sheet option "%s" when "sheets" is defined.', $sheetKey),
                    );
                }
            }

            ConfigPath::assertKnownKeys($config, self::WORKBOOK_ROOT_KEYS, $rootPath);

            foreach ($config['sheets'] as $idx => $sheetConfig) {
                if (!is_array($sheetConfig)) {
                    throw InvalidConfigException::atPath($rootPath->append('sheets')->append($idx), 'Each sheet configuration must be an array.');
                }

                $mergedSheetConfig = array_merge($defaults, $sheetConfig);
                $sheets[] = self::parseSheet($mergedSheetConfig, $rootPath->append('sheets')->append($idx));
            }
        } else {
            // Single-sheet mode: root config contains sheet options
            $allowedSingleSheetKeys = array_merge(self::WORKBOOK_ROOT_KEYS, self::SHEET_KEYS);
            ConfigPath::assertKnownKeys($config, $allowedSingleSheetKeys, $rootPath);

            $mergedSheetConfig = array_merge($defaults, $config);
            $sheets[] = self::parseSheet($mergedSheetConfig, $rootPath);
        }

        if ($formatLower === 'csv' && count($sheets) > 1) {
            throw InvalidConfigException::atPath(
                $rootPath->append('format'),
                'CSV format only supports a single sheet. Multiple sheets were provided.',
            );
        }

        return new WorkbookDefinition(
            format: $formatLower,
            properties: $properties,
            sheets: $sheets,
            defaults: $defaults,
            csv: $csv,
            locale: $locale,
        );
    }

    /**
     * @param array<string, mixed> $sheetConfig
     */
    private static function parseSheet(array $sheetConfig, ConfigPath $path): SheetDefinition
    {
        ConfigPath::assertKnownKeys($sheetConfig, array_merge(self::WORKBOOK_ROOT_KEYS, self::SHEET_KEYS), $path);

        $hasColumns = isset($sheetConfig['columns']) && is_array($sheetConfig['columns']) && $sheetConfig['columns'] !== [];
        $hasHeaders = isset($sheetConfig['headers']) && is_array($sheetConfig['headers']) && $sheetConfig['headers'] !== [];

        if ($hasColumns && $hasHeaders) {
            throw InvalidConfigException::atPath(
                $path,
                'Cannot specify both "columns" and "headers" on the same sheet. Use "columns" with "header" attributes or simple "headers".',
            );
        }

        $startCell = (string) ($sheetConfig['start_cell'] ?? 'A1');
        Coordinate::parseCell($startCell);

        $columns = [];
        if ($hasColumns) {
            /** @var array<array-key, mixed> $rawColumns */
            $rawColumns = $sheetConfig['columns'];
            $columns = self::parseColumns($rawColumns, $path->append('columns'));
        }

        /** @var list<string>|null $headers */
        $headers = null;
        if ($hasHeaders) {
            $headers = array_values(array_map('strval', $sheetConfig['headers']));
        }

        $data = $sheetConfig['data'] ?? [];
        if (!is_iterable($data)) {
            throw InvalidConfigException::atPath($path->append('data'), '"data" must be an iterable (array or Traversable/Generator).');
        }

        /** @var array<string, string>|null $totals */
        $totals = isset($sheetConfig['totals']) && is_array($sheetConfig['totals']) ? $sheetConfig['totals'] : null;

        /** @var list<string> $merges */
        $merges = isset($sheetConfig['merges']) && is_array($sheetConfig['merges'])
            ? array_values(array_map('strval', $sheetConfig['merges']))
            : [];

        $images = [];
        if (isset($sheetConfig['images']) && is_array($sheetConfig['images'])) {
            $images = self::parseImages($sheetConfig['images'], $path->append('images'));
        }

        $tap = $sheetConfig['tap'] ?? null;
        if ($tap !== null && !is_callable($tap)) {
            throw InvalidConfigException::atPath($path->append('tap'), '"tap" must be callable.');
        }

        /** @var list<array<string, mixed>> $conditionalFormatting */
        $conditionalFormatting = isset($sheetConfig['conditional_formatting']) && is_array($sheetConfig['conditional_formatting'])
            ? array_values($sheetConfig['conditional_formatting'])
            : [];

        /** @var bool|null $autoSize */
        $autoSize = array_key_exists('auto_size', $sheetConfig) ? (bool) $sheetConfig['auto_size'] : null;

        return new SheetDefinition(
            name: (string) ($sheetConfig['name'] ?? 'Sheet1'),
            title: isset($sheetConfig['title']) ? (string) $sheetConfig['title'] : null,
            titleStyle: isset($sheetConfig['title_style']) && is_array($sheetConfig['title_style']) ? $sheetConfig['title_style'] : null,
            startCell: $startCell,
            columns: $columns,
            headers: $headers,
            data: $data,
            showHeader: (bool) ($sheetConfig['show_header'] ?? true),
            totals: $totals,
            totalsLabel: isset($sheetConfig['totals_label']) ? (string) $sheetConfig['totals_label'] : null,
            headerStyle: isset($sheetConfig['header_style']) && is_array($sheetConfig['header_style']) ? $sheetConfig['header_style'] : null,
            rowStyle: isset($sheetConfig['row_style']) && is_array($sheetConfig['row_style']) ? $sheetConfig['row_style'] : null,
            zebra: $sheetConfig['zebra'] ?? false,
            borders: $sheetConfig['borders'] ?? false,
            freezeHeader: (bool) ($sheetConfig['freeze_header'] ?? true),
            autoFilter: (bool) ($sheetConfig['auto_filter'] ?? false),
            autoSize: $autoSize,
            autoSizeMaxRows: (int) ($sheetConfig['auto_size_max_rows'] ?? 5000),
            rowHeight: isset($sheetConfig['row_height']) ? (float) $sheetConfig['row_height'] : null,
            merges: $merges,
            images: $images,
            conditionalFormatting: $conditionalFormatting,
            tap: $tap,
        );
    }

    /**
     * @param array<array-key, mixed> $rawColumns
     * @return list<WriteColumn>
     */
    private static function parseColumns(array $rawColumns, ConfigPath $path): array
    {
        $columns = [];
        $isList = array_is_list($rawColumns);

        foreach ($rawColumns as $keyOrIdx => $colDef) {
            $subPath = $path->append((string) $keyOrIdx);

            if (is_string($colDef)) {
                // Shorthand: 'code' => 'Mã NV'
                $key = $isList ? $colDef : (string) $keyOrIdx;
                $header = $isList ? $colDef : $colDef;

                $columns[] = new WriteColumn(
                    key: $key,
                    header: $header,
                    type: 'auto',
                );
                continue;
            }

            if (!is_array($colDef)) {
                throw InvalidConfigException::atPath($subPath, 'Column definition must be a string header or an array.');
            }

            ConfigPath::assertKnownKeys($colDef, self::ALLOWED_COLUMN_KEYS, $subPath);

            $key = $colDef['key'] ?? ($isList ? null : (string) $keyOrIdx);
            if (!is_string($key) || trim($key) === '') {
                throw InvalidConfigException::atPath($subPath->append('key'), 'Column "key" is required.');
            }

            $type = (string) ($colDef['type'] ?? 'auto');
            if ($type === 'formatted') {
                throw InvalidConfigException::atPath(
                    $subPath->append('type'),
                    'Type "formatted" is read-only and cannot be used for writing.',
                );
            }

            if (!TypeRegistry::has($type)) {
                throw InvalidConfigException::atPath(
                    $subPath->append('type'),
                    sprintf('Unknown column type "%s".', $type),
                );
            }

            $valueCallback = $colDef['value'] ?? null;
            if ($valueCallback !== null && !is_callable($valueCallback)) {
                throw InvalidConfigException::atPath($subPath->append('value'), '"value" must be callable.');
            }

            $columns[] = new WriteColumn(
                key: trim($key),
                header: isset($colDef['header']) ? (string) $colDef['header'] : null,
                type: $type,
                format: isset($colDef['format']) ? (string) $colDef['format'] : null,
                width: $colDef['width'] ?? null,
                style: self::resolveColumnStyle($colDef),
                value: $valueCallback,
                default: $colDef['default'] ?? null,
                labels: isset($colDef['labels']) && is_array($colDef['labels']) ? $colDef['labels'] : null,
                align: isset($colDef['align']) && is_string($colDef['align']) ? strtolower(trim($colDef['align'])) : null,
                wrap: isset($colDef['wrap']) ? (bool) $colDef['wrap'] : null,
                validation: self::resolveValidation($colDef),
                conditional: isset($colDef['conditional']) && is_array($colDef['conditional']) ? $colDef['conditional'] : null,
                options: $colDef,
            );
        }

        return $columns;
    }
    /**
     * @param array<string, mixed> $colDef
     * @return array<string, mixed>|null
     */
    private static function resolveColumnStyle(array $colDef): ?array
    {
        $style = isset($colDef['style']) && is_array($colDef['style']) ? $colDef['style'] : [];
        if (isset($colDef['align']) && is_string($colDef['align']) && !isset($style['align'])) {
            $style['align'] = strtolower(trim($colDef['align']));
        }
        if (isset($colDef['wrap']) && !isset($style['wrap'])) {
            $style['wrap'] = (bool) $colDef['wrap'];
        }

        return $style !== [] ? $style : null;
    }
    /**
     * @param array<string, mixed> $colDef
     * @return array<string, mixed>|null
     */
    private static function resolveValidation(array $colDef): ?array
    {
        if (!isset($colDef['validation'])) {
            return null;
        }

        $val = $colDef['validation'];
        if (is_string($val)) {
            return [
                'type' => 'list',
                'options' => array_map('trim', explode(',', $val)),
            ];
        }

        if (is_array($val)) {
            if (array_is_list($val)) {
                return [
                    'type' => 'list',
                    'options' => $val,
                ];
            }

            return $val;
        }

        return null;
    }
    /**
     * @param array<array-key, mixed> $rawImages
     * @return list<ImageDefinition>
     */
    private static function parseImages(array $rawImages, ConfigPath $path): array
    {
        $images = [];

        // Single image definition: ['path' => '...']
        if (isset($rawImages['path']) && is_string($rawImages['path'])) {
            $rawImages = [$rawImages];
        }

        foreach ($rawImages as $idx => $imgDef) {
            $subPath = $path->append((string) $idx);
            if (is_string($imgDef)) {
                $imgDef = ['path' => $imgDef];
            }
            if (!is_array($imgDef)) {
                throw InvalidConfigException::atPath($subPath, 'Image definition must be a string path or an array.');
            }

            $imgPath = (string) ($imgDef['path'] ?? '');
            if ($imgPath === '') {
                throw InvalidConfigException::atPath($subPath->append('path'), 'Image "path" is required.');
            }
            if (!file_exists($imgPath)) {
                throw InvalidConfigException::atPath($subPath->append('path'), sprintf('Image file "%s" does not exist.', $imgPath));
            }

            $cell = (string) ($imgDef['cell'] ?? $imgDef['coordinates'] ?? 'A1');
            Coordinate::parseCell($cell);

            $images[] = new ImageDefinition(
                path: $imgPath,
                cell: strtoupper($cell),
                width: isset($imgDef['width']) ? (int) $imgDef['width'] : null,
                height: isset($imgDef['height']) ? (int) $imgDef['height'] : null,
                offsetX: (int) ($imgDef['offset_x'] ?? 0),
                offsetY: (int) ($imgDef['offset_y'] ?? 0),
                name: isset($imgDef['name']) ? (string) $imgDef['name'] : null,
                description: isset($imgDef['description']) ? (string) $imgDef['description'] : null,
            );
        }

        return $images;
    }
}
