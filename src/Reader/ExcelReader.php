<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use ExcelHelper\Config\ReadConfigParser;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\RowValidationException;
use ExcelHelper\Exceptions\SheetNotFoundException;
use ExcelHelper\Support\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\BaseReader;
use PhpOffice\PhpSpreadsheet\Reader\Csv as CsvReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Throwable;

/**
 * Core reader coordinating file inspection, PhpSpreadsheet loading, column resolution, and row mapping.
 */
final class ExcelReader
{
    /**
     * @param array<string, mixed> $options
     *
     * @throws ReaderException
     * @throws SheetNotFoundException
     * @throws RowValidationException
     */
    public static function read(string $path, array $options = []): ReadResult
    {
        $opts = ReadConfigParser::parse($options);

        $inspection = FileInspector::inspect($path, $opts);
        $readerType = $inspection['readerType'];

        /** @var BaseReader $reader */
        $reader = IOFactory::createReader($readerType);

        // Check if any column needs formatted display value
        $needsFormatted = false;
        foreach ($opts->columns as $col) {
            if ($col->type === 'formatted') {
                $needsFormatted = true;
                break;
            }
        }

        $reader->setReadDataOnly(!$needsFormatted);
        $reader->setReadEmptyCells(false);

        // Configure CSV specific options
        if ($reader instanceof CsvReader) {
            self::configureCsvReader($reader, $opts);
        }

        // Worksheet selection and safety checks
        $sheetNames = [];
        try {
            $sheetNames = array_values($reader->listWorksheetNames($path));
        } catch (Throwable) {
            // Some readers might fail to list worksheet names on empty/unusual files
        }

        $targetSheetName = self::resolveTargetSheetName($opts->sheet, $sheetNames);

        if ($targetSheetName !== null) {
            $reader->setLoadSheetsOnly([$targetSheetName]);
        }

        // Check max_rows safety limit before loading full sheet
        if ($opts->maxRows !== null) {
            self::checkMaxRows($reader, $path, $targetSheetName, $opts->maxRows);
        }

        // Set chunk read filter to minimize loaded memory
        $allowedColLetters = null;
        if ($opts->columns !== []) {
            $colLetters = [];
            foreach ($opts->columns as $c) {
                if ($c->column !== null) {
                    $colLetters[] = Coordinate::columnLetter($c->column);
                }
            }
            if ($colLetters !== [] && count($colLetters) === count($opts->columns)) {
                $allowedColLetters = $colLetters;
            }
        }

        $filter = new ChunkReadFilter($opts->startRow, $opts->endRow, $opts->headerRow, $allowedColLetters);
        $reader->setReadFilter($filter);

        /** @var ?Spreadsheet $spreadsheet */
        $spreadsheet = null;

        try {
            $spreadsheet = $reader->load($path);
        } catch (Throwable $e) {
            throw ReaderException::cannotOpenFile($path, $e);
        }

        try {
            /** @var Worksheet $worksheet */
            $worksheet = $targetSheetName !== null
                ? ($spreadsheet->getSheetByName($targetSheetName) ?? $spreadsheet->getActiveSheet())
                : $spreadsheet->getActiveSheet();

            $resolvedColumns = ColumnResolver::resolve($worksheet, $opts);

            // Extract headers if headerRow is configured
            $headers = [];
            if ($opts->headerRow !== null) {
                $highestCol = $worksheet->getHighestDataColumn();
                $maxColIdx = Coordinate::columnIndex($highestCol);
                for ($c = 1; $c <= $maxColIdx; $c++) {
                    $letter = Coordinate::columnLetter($c);
                    $cellVal = $worksheet->getCell($letter . $opts->headerRow)->getValue();
                    if ($cellVal !== null && $cellVal !== '') {
                        $headers[$letter] = (string) $cellVal;
                    }
                }
            }

            $rows = [];
            $allErrors = [];

            $highestDataRow = $worksheet->getHighestDataRow();
            $endRow = $opts->endRow !== null ? min($opts->endRow, $highestDataRow) : $highestDataRow;

            $globalOptions = [
                'date_format' => $opts->dateFormat,
                'datetime_format' => $opts->datetimeFormat,
                'locale' => $opts->locale,
                'trim' => $opts->trim,
            ];

            for ($r = $opts->startRow; $r <= $endRow; $r++) {
                $rawRow = [];

                foreach ($resolvedColumns as $col) {
                    if ($col->column === null) {
                        continue;
                    }
                    $colIndex = (int) $col->column;
                    $cellAddress = Coordinate::formatCell($colIndex, $r);
                    $cell = $worksheet->getCell($cellAddress);

                    if ($col->type === 'formatted') {
                        $val = $cell->getFormattedValue();
                    } elseif ($opts->calculateFormulas) {
                        try {
                            $val = $cell->getCalculatedValue();
                        } catch (Throwable $e) {
                            $val = $cell->getValue();
                        }
                    } else {
                        $val = $cell->getValue();
                    }

                    $rawRow[$colIndex] = $val;
                }

                if (RowMapper::isRowEmpty($rawRow, $resolvedColumns)) {
                    if ($opts->stopAtEmptyRow) {
                        break;
                    }
                    if ($opts->skipEmptyRows) {
                        continue;
                    }
                }

                $mapped = RowMapper::map($r, $rawRow, $resolvedColumns, $opts, $globalOptions);

                if ($mapped['errors'] !== []) {
                    if ($opts->onError === 'throw') {
                        throw RowValidationException::fromRowError($mapped['errors'][0]);
                    }

                    foreach ($mapped['errors'] as $err) {
                        $allErrors[] = $err;
                    }

                    if ($opts->onError === 'skip_row') {
                        continue;
                    }
                }

                $rows[] = $mapped['row'];
            }

            return new ReadResult(
                rows: $rows,
                errors: $allErrors,
                headers: $headers,
                sheetName: $worksheet->getTitle(),
            );
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    private static function configureCsvReader(CsvReader $reader, ReadOptions $options): void
    {
        $csv = $options->csv;

        if (isset($csv['delimiter']) && is_string($csv['delimiter']) && $csv['delimiter'] !== 'auto') {
            $reader->setDelimiter($csv['delimiter']);
        }

        if (isset($csv['enclosure']) && is_string($csv['enclosure'])) {
            $reader->setEnclosure($csv['enclosure']);
        }

        if (isset($csv['escape']) && is_string($csv['escape'])) {
            $reader->setEscapeCharacter($csv['escape']);
        }

        if (isset($csv['input_encoding']) && is_string($csv['input_encoding']) && $csv['input_encoding'] !== 'auto') {
            $reader->setInputEncoding($csv['input_encoding']);
        }
    }

    /**
     * @param list<string> $availableSheets
     */
    private static function resolveTargetSheetName(int|string $sheet, array $availableSheets): ?string
    {
        if ($availableSheets === []) {
            return is_string($sheet) ? $sheet : null;
        }

        if (is_int($sheet)) {
            if (!isset($availableSheets[$sheet])) {
                throw SheetNotFoundException::byIndex($sheet, $availableSheets);
            }

            return $availableSheets[$sheet];
        }

        if (!in_array($sheet, $availableSheets, true)) {
            throw SheetNotFoundException::byName($sheet, $availableSheets);
        }

        return $sheet;
    }

    private static function checkMaxRows(BaseReader $reader, string $path, ?string $sheetName, int $maxRows): void
    {
        try {
            $info = $reader->listWorksheetInfo($path);
        } catch (Throwable) {
            return;
        }

        foreach ($info as $sheetInfo) {
            if ($sheetName === null || $sheetInfo['worksheetName'] === $sheetName) {
                $totalRows = (int) $sheetInfo['totalRows'];
                if ($totalRows > $maxRows) {
                    throw ReaderException::maxRowsExceeded($maxRows, $totalRows);
                }
            }
        }
    }
    /**
     * Read a spreadsheet file in chunks.
     *
     * @param array<string, mixed> $options
     * @param callable(list<array<string, mixed>>, ChunkContext): (bool|void) $callback
     */
    public static function chunk(
        string $path,
        array $options,
        int $chunkSize,
        callable $callback,
    ): ReadSummary {
        if ($chunkSize < 1) {
            throw InvalidConfigException::withMessage('Chunk size must be >= 1.');
        }

        $opts = ReadConfigParser::parse($options);
        $inspection = FileInspector::inspect($path, $opts);
        $readerType = $inspection['readerType'];

        if ($inspection['format'] === 'csv') {
            return CsvStreamReader::chunk($path, $opts, $chunkSize, $callback);
        }

        /** @var BaseReader $reader */
        $reader = IOFactory::createReader($readerType);

        $needsFormatted = false;
        foreach ($opts->columns as $col) {
            if ($col->type === 'formatted') {
                $needsFormatted = true;
                break;
            }
        }
        $reader->setReadDataOnly(!$needsFormatted);
        $reader->setReadEmptyCells(false);

        $sheetNames = array_values($reader->listWorksheetNames($path));
        $targetSheetName = self::resolveTargetSheetName($opts->sheet, $sheetNames);

        if ($targetSheetName !== null) {
            $reader->setLoadSheetsOnly([$targetSheetName]);
        }

        // Determine total rows
        $info = $reader->listWorksheetInfo($path);
        $sheetTotalRows = 0;
        foreach ($info as $sInfo) {
            if ($targetSheetName === null || $sInfo['worksheetName'] === $targetSheetName) {
                $sheetTotalRows = (int) $sInfo['totalRows'];
                break;
            }
        }

        if ($opts->maxRows !== null && $sheetTotalRows > $opts->maxRows) {
            throw ReaderException::maxRowsExceeded($opts->maxRows, $sheetTotalRows);
        }

        $startRow = $opts->startRow;
        $maxEndRow = $opts->endRow !== null ? min($opts->endRow, $sheetTotalRows) : $sheetTotalRows;

        $globalOptions = [
            'date_format' => $opts->dateFormat,
            'datetime_format' => $opts->datetimeFormat,
            'locale' => $opts->locale,
            'trim' => $opts->trim,
        ];

        $filter = new ChunkReadFilter($startRow, min($startRow + $chunkSize - 1, $maxEndRow), $opts->headerRow);
        $reader->setReadFilter($filter);

        $headers = [];
        $resolvedColumns = null;
        $allErrors = [];
        $totalProcessed = 0;
        $chunkIndex = 0;
        $stop = false;

        for ($currStart = $startRow; $currStart <= $maxEndRow; $currStart += $chunkSize) {
            $currEnd = min($currStart + $chunkSize - 1, $maxEndRow);
            $filter->setRows($currStart, $currEnd);

            /** @var Spreadsheet $spreadsheet */
            $spreadsheet = $reader->load($path);

            try {
                /** @var Worksheet $worksheet */
                $worksheet = $targetSheetName !== null
                    ? ($spreadsheet->getSheetByName($targetSheetName) ?? $spreadsheet->getActiveSheet())
                    : $spreadsheet->getActiveSheet();

                if ($resolvedColumns === null) {
                    $resolvedColumns = ColumnResolver::resolve($worksheet, $opts);

                    if ($opts->headerRow !== null) {
                        $highestCol = $worksheet->getHighestDataColumn();
                        $maxColIdx = Coordinate::columnIndex($highestCol);
                        for ($c = 1; $c <= $maxColIdx; $c++) {
                            $letter = Coordinate::columnLetter($c);
                            $val = $worksheet->getCell($letter . $opts->headerRow)->getValue();
                            if ($val !== null && $val !== '') {
                                $headers[$letter] = (string) $val;
                            }
                        }
                    }
                }

                $chunkRows = [];
                $chunkErrors = [];

                for ($r = $currStart; $r <= $currEnd; $r++) {
                    $rawRow = [];
                    foreach ($resolvedColumns as $col) {
                        if ($col->column === null) {
                            continue;
                        }
                        $colIndex = (int) $col->column;
                        $cellAddress = Coordinate::formatCell($colIndex, $r);
                        $cell = $worksheet->getCell($cellAddress);
                        $val = $col->type === 'formatted'
                            ? $cell->getFormattedValue()
                            : ($opts->calculateFormulas ? $cell->getCalculatedValue() : $cell->getValue());
                        $rawRow[$colIndex] = $val;
                    }

                    if (RowMapper::isRowEmpty($rawRow, $resolvedColumns)) {
                        if ($opts->stopAtEmptyRow) {
                            $stop = true;
                            break;
                        }
                        if ($opts->skipEmptyRows) {
                            continue;
                        }
                    }

                    $mapped = RowMapper::map($r, $rawRow, $resolvedColumns, $opts, $globalOptions);
                    if ($mapped['errors'] !== []) {
                        if ($opts->onError === 'throw') {
                            throw RowValidationException::fromRowError($mapped['errors'][0]);
                        }
                        foreach ($mapped['errors'] as $err) {
                            $chunkErrors[] = $err;
                            $allErrors[] = $err;
                        }
                        if ($opts->onError === 'skip_row') {
                            continue;
                        }
                    }

                    $chunkRows[] = $mapped['row'];
                    $totalProcessed++;
                }

                if ($chunkRows !== []) {
                    $chunkContext = new ChunkContext(
                        chunkIndex: $chunkIndex,
                        startRow: $currStart,
                        endRow: $currEnd,
                        rowCount: count($chunkRows),
                        totalRowsProcessed: $totalProcessed,
                        errors: $chunkErrors,
                    );

                    $ret = $callback($chunkRows, $chunkContext);
                    $chunkIndex++;

                    if ($ret === false || $chunkContext->shouldStop()) {
                        $stop = true;
                    }
                }

                if ($stop) {
                    break;
                }
            } finally {
                $spreadsheet->disconnectWorksheets();
                unset($spreadsheet);
            }
        }

        return new ReadSummary(
            totalRows: $totalProcessed,
            totalChunks: $chunkIndex,
            errors: $allErrors,
            headers: $headers,
            sheetName: $targetSheetName ?? 'Worksheet',
        );
    }

    /**
     * @return list<string>
     */
    public static function sheetNames(string $path): array
    {
        $opts = ReadConfigParser::parse([]);
        $inspection = FileInspector::inspect($path, $opts);
        /** @var BaseReader $reader */
        $reader = IOFactory::createReader($inspection['readerType']);

        return array_values($reader->listWorksheetNames($path));
    }

    /**
     * @return list<SheetInfo>
     */
    public static function info(string $path): array
    {
        $opts = ReadConfigParser::parse([]);
        $inspection = FileInspector::inspect($path, $opts);
        /** @var BaseReader $reader */
        $reader = IOFactory::createReader($inspection['readerType']);

        $rawInfo = $reader->listWorksheetInfo($path);
        $result = [];

        foreach ($rawInfo as $item) {
            $lastColIndex = (int) $item['lastColumnIndex'];
            $result[] = new SheetInfo(
                name: (string) $item['worksheetName'],
                totalRows: (int) $item['totalRows'],
                totalColumns: (int) $item['totalColumns'],
                lastColumnLetter: (string) $item['lastColumnLetter'],
                lastColumnIndex: $lastColIndex,
            );
        }

        return $result;
    }

    /**
     * @return array<string, string>
     */
    public static function headers(string $path, int $row = 1, int|string $sheet = 0): array
    {
        $res = self::read($path, [
            'sheet' => $sheet,
            'header_row' => $row,
            'start_row' => $row + 1,
            'end_row' => $row + 1,
        ]);

        return $res->headers();
    }
}
