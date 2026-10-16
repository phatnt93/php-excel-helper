<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\RowValidationException;
use ExcelHelper\Exceptions\SheetNotFoundException;
use ExcelHelper\Support\Coordinate;
use Generator;
use XMLReader;
use ZipArchive;

/**
 * Fast streaming XLSX reader using PHP's native XMLReader over zip archive.
 * Provides true O(1) memory consumption for ultra-large XLSX spreadsheets.
 */
final class FastXlsxStreamReader
{
    /**
     * Stream rows as a generator.
     *
     * @return Generator<int, array<string, mixed>>
     *
     * @throws ReaderException
     * @throws RowValidationException
     * @throws SheetNotFoundException
     */
    public static function stream(string $path, ReadOptions $options): Generator
    {
        $realPath = realpath($path);
        if ($realPath === false || !is_file($realPath)) {
            throw ReaderException::fileNotFound($path);
        }

        $sharedStrings = self::loadSharedStrings($realPath);
        $sheetXmlPath = self::resolveSheetXml($realPath, $options->sheet);

        $reader = new XMLReader();
        if (!$reader->open('zip://' . $realPath . '#' . $sheetXmlPath)) {
            throw ReaderException::cannotOpenFile($path);
        }

        $resolvedColumns = null;
        $headerLookup = [];
        $lastRowNumber = null;
        $headerRow = $options->headerRow;
        $startRow = $options->startRow;
        $endRow = $options->endRow;
        $globalOptions = [
            'date_format' => $options->dateFormat,
            'datetime_format' => $options->datetimeFormat,
            'locale' => $options->locale,
            'trim' => $options->trim,
        ];

        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || $reader->name !== 'row') {
                    continue;
                }

                $rowNumber = (int) ($reader->getAttribute('r') ?? '0');
                if ($rowNumber === 0) {
                    continue;
                }

                // Read all cells for this row
                $rawRow = self::parseRowCells($reader, $sharedStrings, $rowNumber);

                // Handle header row detection
                if ($headerRow !== null && $rowNumber === $headerRow) {
                    foreach ($rawRow as $cIdx => $val) {
                        $headerLookup[$cIdx] = is_string($val) ? $val : (string) $val;
                    }
                    continue;
                }

                // Skip rows before startRow
                if ($rowNumber < $startRow) {
                    continue;
                }

                // Stop if beyond endRow
                if ($endRow !== null && $rowNumber > $endRow) {
                    break;
                }

                if ($options->stopAtEmptyRow && $lastRowNumber !== null && $rowNumber > $lastRowNumber + 1) {
                    break;
                }
                $lastRowNumber = $rowNumber;

                // Resolve columns on first data row
                if ($resolvedColumns === null) {
                    $resolvedColumns = self::resolveColumns($options, $headerLookup, $rawRow);
                }

                // Check empty row
                if (RowMapper::isRowEmpty($rawRow, $resolvedColumns)) {
                    if ($options->stopAtEmptyRow) {
                        break;
                    }
                    if ($options->skipEmptyRows) {
                        continue;
                    }
                }

                // Map row
                $mapped = RowMapper::map($rowNumber, $rawRow, $resolvedColumns, $options, $globalOptions);

                if ($mapped['errors'] !== []) {
                    if ($options->onError === 'throw') {
                        throw RowValidationException::fromRowError($mapped['errors'][0]);
                    }
                    if ($options->onError === 'skip_row') {
                        continue;
                    }
                }

                yield $rowNumber => $mapped['row'];
            }
        } finally {
            $reader->close();
        }
    }

    /**
     * Read entire XLSX via fast XML streaming.
     *
     * @throws ReaderException
     * @throws RowValidationException
     * @throws SheetNotFoundException
     */
    public static function read(string $path, ReadOptions $options): ReadResult
    {
        $realPath = realpath($path);
        if ($realPath === false || !is_file($realPath)) {
            throw ReaderException::fileNotFound($path);
        }

        $sharedStrings = self::loadSharedStrings($realPath);
        $sheetXmlPath = self::resolveSheetXml($realPath, $options->sheet);

        $reader = new XMLReader();
        if (!$reader->open('zip://' . $realPath . '#' . $sheetXmlPath)) {
            throw ReaderException::cannotOpenFile($path);
        }

        $resolvedColumns = null;
        $headerLookup = [];
        $headersResult = [];
        $rows = [];
        $errors = [];
        $lastRowNumber = null;
        $headerRow = $options->headerRow;
        $startRow = $options->startRow;
        $endRow = $options->endRow;
        $globalOptions = [
            'date_format' => $options->dateFormat,
            'datetime_format' => $options->datetimeFormat,
            'locale' => $options->locale,
            'trim' => $options->trim,
        ];

        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || $reader->name !== 'row') {
                    continue;
                }

                $rowNumber = (int) ($reader->getAttribute('r') ?? '0');
                if ($rowNumber === 0) {
                    continue;
                }

                $rawRow = self::parseRowCells($reader, $sharedStrings, $rowNumber);

                if ($headerRow !== null && $rowNumber === $headerRow) {
                    foreach ($rawRow as $cIdx => $val) {
                        $colLetter = Coordinate::columnLetter($cIdx);
                        $headerText = is_string($val) ? $val : (string) $val;
                        $headerLookup[$cIdx] = $headerText;
                        $headersResult[$colLetter] = $headerText;
                    }
                    continue;
                }

                if ($rowNumber < $startRow) {
                    continue;
                }

                if ($endRow !== null && $rowNumber > $endRow) {
                    break;
                }

                if ($options->stopAtEmptyRow && $lastRowNumber !== null && $rowNumber > $lastRowNumber + 1) {
                    break;
                }
                $lastRowNumber = $rowNumber;

                if ($resolvedColumns === null) {
                    $resolvedColumns = self::resolveColumns($options, $headerLookup, $rawRow);
                }

                if (RowMapper::isRowEmpty($rawRow, $resolvedColumns)) {
                    if ($options->stopAtEmptyRow) {
                        break;
                    }
                    if ($options->skipEmptyRows) {
                        continue;
                    }
                }

                $mapped = RowMapper::map($rowNumber, $rawRow, $resolvedColumns, $options, $globalOptions);

                if ($mapped['errors'] !== []) {
                    if ($options->onError === 'throw') {
                        throw RowValidationException::fromRowError($mapped['errors'][0]);
                    }
                    foreach ($mapped['errors'] as $err) {
                        $errors[] = $err;
                    }
                    if ($options->onError === 'skip_row') {
                        continue;
                    }
                }

                $rows[] = $mapped['row'];
            }
        } finally {
            $reader->close();
        }

        return new ReadResult(
            rows: $rows,
            errors: $errors,
            headers: $headersResult,
            sheetName: is_string($options->sheet) ? $options->sheet : 'Sheet1',
        );
    }

    /**
     * @return list<string>
     */
    private static function loadSharedStrings(string $zipPath): array
    {
        $reader = new XMLReader();
        if (!@$reader->open('zip://' . $zipPath . '#xl/sharedStrings.xml')) {
            return [];
        }

        $strings = [];
        try {
            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::ELEMENT && $reader->name === 'si') {
                    $xml = $reader->readOuterXml();
                    if (preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $xml, $m)) {
                        $strings[] = html_entity_decode(implode('', $m[1]), ENT_QUOTES | ENT_XML1, 'UTF-8');
                    } else {
                        $strings[] = '';
                    }
                }
            }
        } finally {
            $reader->close();
        }

        return $strings;
    }

    /**
     * @throws SheetNotFoundException
     */
    private static function resolveSheetXml(string $zipPath, int|string $sheet): string
    {
        // Check workbook.xml for sheet mappings
        $zip = new ZipArchive();
        if ($zip->open($zipPath) !== true) {
            return 'xl/worksheets/sheet1.xml';
        }

        $sheets = [];
        $xml = $zip->getFromName('xl/workbook.xml');
        $zip->close();

        if ($xml !== false && preg_match_all('/<sheet[^>]+name="([^"]+)"[^>]+sheetId="([^"]+)"/u', $xml, $m)) {
            foreach ($m[1] as $idx => $name) {
                $sheets[] = [
                    'index' => $idx,
                    'name' => $name,
                    'file' => 'xl/worksheets/sheet' . ($idx + 1) . '.xml',
                ];
            }
        }

        if ($sheets === []) {
            return 'xl/worksheets/sheet1.xml';
        }

        if (is_int($sheet)) {
            if (isset($sheets[$sheet])) {
                return $sheets[$sheet]['file'];
            }
            $names = array_column($sheets, 'name');
            throw SheetNotFoundException::byIndex($sheet, $names);
        }

        foreach ($sheets as $s) {
            if ($s['name'] === $sheet) {
                return $s['file'];
            }
        }

        $names = array_column($sheets, 'name');
        throw SheetNotFoundException::byName($sheet, $names);
    }

    /**
     * @param list<string> $sharedStrings
     * @return array<int, mixed>
     */
    private static function parseRowCells(XMLReader $reader, array $sharedStrings, int $rowNumber): array
    {
        $rawRow = [];
        $rowXml = $reader->readOuterXml();

        if (preg_match_all('/<c\s+([^>]+)>(?:<f[^>]*>.*?<\/f>)?(?:<v>(.*?)<\/v>|<is>(?:<t[^>]*>(.*?)<\/t>)<\/is>)?<\/c>|<c\s+([^>]+)\/>/s', $rowXml, $matches, PREG_SET_ORDER)) {
            foreach ($matches as $match) {
                $match1 = (string) ($match[1] ?? '');
                $match4 = (string) ($match[4] ?? '');
                $attrs = $match1 !== '' ? $match1 : $match4;
                $vVal = $match[2] ?? null;
                $isVal = $match[3] ?? null;

                preg_match('/r="([A-Z]+)(\d+)"/', $attrs, $rMatch);
                $colLetter = $rMatch[1] ?? 'A';
                $cIdx = Coordinate::columnIndex($colLetter);

                preg_match('/t="([^"]+)"/', $attrs, $tMatch);
                $type = $tMatch[1] ?? '';

                $cellVal = null;
                if ($type === 's' && $vVal !== null) {
                    $sIdx = (int) $vVal;
                    $cellVal = $sharedStrings[$sIdx] ?? '';
                } elseif ($type === 'inlineStr' && $isVal !== null) {
                    $cellVal = html_entity_decode($isVal, ENT_QUOTES | ENT_XML1, 'UTF-8');
                } elseif ($type === 'b') {
                    $cellVal = ($vVal === '1');
                } elseif ($vVal !== null && $vVal !== '') {
                    $cellVal = is_numeric($vVal) ? (str_contains((string) $vVal, '.') ? (float) $vVal : (int) $vVal) : $vVal;
                }

                $rawRow[$cIdx] = $cellVal;
            }
        }

        return $rawRow;
    }

    /**
     * @param array<int, string> $headerLookup
     * @param array<int, mixed> $rawRow
     * @return list<ReadColumn>
     */
    private static function resolveColumns(ReadOptions $options, array $headerLookup, array $rawRow): array
    {
        if ($options->columns !== []) {
            $resolved = [];
            foreach ($options->columns as $col) {
                if ($col->header !== null && $headerLookup !== []) {
                    $matchedIdx = null;
                    $headersToMatch = is_array($col->header) ? $col->header : [$col->header];
                    foreach ($headerLookup as $idx => $headerText) {
                        foreach ($headersToMatch as $header) {
                            if (mb_strtolower(trim($headerText)) === mb_strtolower(trim($header))) {
                                $matchedIdx = $idx;
                                break 2;
                            }
                        }
                    }

                    if ($matchedIdx !== null) {
                        $resolved[] = $col->withResolvedColumn($matchedIdx);
                    } elseif ($col->optional) {
                        $resolved[] = $col->withResolvedColumn(null);
                    } else {
                        throw ReaderException::missingHeaderColumn(is_array($col->header) ? implode(' / ', $col->header) : $col->header);
                    }
                } else {
                    $resolved[] = $col;
                }
            }

            return $resolved;
        }

        // Auto-detect columns from rawRow or headerLookup
        $cols = [];
        $keys = $headerLookup !== [] ? array_keys($headerLookup) : array_keys($rawRow);
        sort($keys);

        foreach ($keys as $cIdx) {
            $keyName = $headerLookup[$cIdx] ?? Coordinate::columnLetter($cIdx);
            $cols[] = new ReadColumn(
                key: $keyName,
                column: $cIdx,
                header: $headerLookup[$cIdx] ?? null,
                type: 'auto',
            );
        }

        return $cols;
    }
}
