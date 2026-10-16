<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\RowValidationException;
use ExcelHelper\Support\Coordinate;
use SplFileObject;

/**
 * High-performance streaming CSV reader with auto-delimiter and auto-encoding detection.
 */
final class CsvStreamReader
{
    private const DELIMITER_CANDIDATES = [',', ';', "\t", '|'];

    /**
     * @return \Generator<int, array<string, mixed>>
     *
     * @throws ReaderException
     * @throws RowValidationException
     */
    public static function stream(string $path, ReadOptions $options): \Generator
    {
        $delimiter = self::resolveDelimiter($path, $options->csv['delimiter'] ?? 'auto');
        $enclosure = (string) ($options->csv['enclosure'] ?? '"');
        $escape = (string) ($options->csv['escape'] ?? '\\');
        $encoding = self::resolveEncoding($path, $options->csv['input_encoding'] ?? 'auto');

        if ($encoding !== 'UTF-8') {
            $rawContent = file_get_contents($path);
            if ($rawContent !== false) {
                $utf8Content = mb_convert_encoding($rawContent, 'UTF-8', $encoding);
                $file = new SplFileObject('php://temp', 'r+');
                if (is_string($utf8Content)) {
                    $file->fwrite($utf8Content);
                }
                $file->rewind();
            } else {
                $file = new SplFileObject($path, 'r');
            }
        } else {
            $file = new SplFileObject($path, 'r');
        }

        $file->setFlags(SplFileObject::READ_CSV | SplFileObject::DROP_NEW_LINE | SplFileObject::SKIP_EMPTY);
        $file->setCsvControl($delimiter, $enclosure, $escape);

        $lineNum = 0;
        $headers = [];
        $headerLookup = null;
        $resolvedColumns = [];
        $columnsInitialized = false;

        $globalOptions = [
            'date_format' => $options->dateFormat,
            'datetime_format' => $options->datetimeFormat,
            'locale' => $options->locale,
            'trim' => $options->trim,
        ];

        while (!$file->eof()) {
            /** @var list<mixed>|false|null $rawRow */
            $rawRow = $file->fgetcsv($delimiter, $enclosure, $escape);

            if ($rawRow === false || $rawRow === null || $rawRow === [null]) {
                continue;
            }

            $lineNum++;

            // Convert encoding and strip BOM on line 1
            if ($lineNum === 1) {
                $rawRow = self::stripBomFromRow($rawRow);
            }

            // Convert to 1-based indexed raw row
            $indexedRow = [];
            foreach ($rawRow as $colIdx => $colVal) {
                $indexedRow[$colIdx + 1] = $colVal;
            }

            // Handle header row
            if ($options->headerRow !== null && $lineNum === $options->headerRow) {
                foreach ($indexedRow as $cIdx => $cVal) {
                    $letter = Coordinate::columnLetter($cIdx);
                    $headers[$letter] = trim((string) $cVal);
                }
                $headerLookup = self::buildHeaderLookup($headers);
                continue;
            }

            if ($lineNum < $options->startRow) {
                continue;
            }

            if ($options->endRow !== null && $lineNum > $options->endRow) {
                break;
            }

            // Initialize columns on first data row
            if (!$columnsInitialized) {
                $resolvedColumns = self::resolveColumns($indexedRow, $options, $headerLookup);
                $columnsInitialized = true;
            }

            // Check empty row
            if (RowMapper::isRowEmpty($indexedRow, $resolvedColumns)) {
                if ($options->stopAtEmptyRow) {
                    break;
                }
                if ($options->skipEmptyRows) {
                    continue;
                }
            }

            $mapped = RowMapper::map($lineNum, $indexedRow, $resolvedColumns, $options, $globalOptions);

            if ($mapped['errors'] !== []) {
                if ($options->onError === 'throw') {
                    throw RowValidationException::fromRowError($mapped['errors'][0]);
                }

                if ($options->onError === 'skip_row') {
                    continue;
                }
            }

            yield $lineNum => $mapped['row'];
        }
    }
    public static function chunk(
        string $path,
        ReadOptions $options,
        int $chunkSize,
        callable $callback,
    ): ReadSummary {
        $delimiter = self::resolveDelimiter($path, $options->csv['delimiter'] ?? 'auto');
        $enclosure = (string) ($options->csv['enclosure'] ?? '"');
        $escape = (string) ($options->csv['escape'] ?? '\\');
        $encoding = self::resolveEncoding($path, $options->csv['input_encoding'] ?? 'auto');

        if ($encoding !== 'UTF-8') {
            $rawContent = file_get_contents($path);
            if ($rawContent !== false) {
                $utf8Content = mb_convert_encoding($rawContent, 'UTF-8', $encoding);
                $file = new SplFileObject('php://temp', 'r+');
                if (is_string($utf8Content)) {
                    $file->fwrite($utf8Content);
                }
                $file->rewind();
            } else {
                $file = new SplFileObject($path, 'r');
            }
        } else {
            $file = new SplFileObject($path, 'r');
        }

        $file->setFlags(SplFileObject::READ_CSV | SplFileObject::DROP_NEW_LINE | SplFileObject::SKIP_EMPTY);
        $file->setCsvControl($delimiter, $enclosure, $escape);

        $lineNum = 0;
        $headers = [];
        $headerLookup = null;
        $resolvedColumns = [];
        $columnsInitialized = false;

        $globalOptions = [
            'date_format' => $options->dateFormat,
            'datetime_format' => $options->datetimeFormat,
            'locale' => $options->locale,
            'trim' => $options->trim,
        ];

        $currentChunk = [];
        $currentChunkErrors = [];
        $allErrors = [];
        $chunkIndex = 0;
        $totalRowsProcessed = 0;
        $stopStreaming = false;

        while (!$file->eof()) {
            /** @var list<mixed>|false|null $rawRow */
            $rawRow = $file->fgetcsv($delimiter, $enclosure, $escape);

            if ($rawRow === false || $rawRow === null || $rawRow === [null]) {
                continue;
            }

            $lineNum++;

            // Convert encoding and strip BOM on line 1
            if ($lineNum === 1) {
                $rawRow = self::stripBomFromRow($rawRow);
            }


            // Convert to 1-based indexed raw row
            $indexedRow = [];
            foreach ($rawRow as $colIdx => $colVal) {
                $indexedRow[$colIdx + 1] = $colVal;
            }

            // Handle header row
            if ($options->headerRow !== null && $lineNum === $options->headerRow) {
                foreach ($indexedRow as $cIdx => $cVal) {
                    $letter = Coordinate::columnLetter($cIdx);
                    $headers[$letter] = trim((string) $cVal);
                }
                $headerLookup = self::buildHeaderLookup($headers);
                continue;
            }

            if ($lineNum < $options->startRow) {
                continue;
            }

            if ($options->endRow !== null && $lineNum > $options->endRow) {
                break;
            }

            // Initialize columns on first data row
            if (!$columnsInitialized) {
                $resolvedColumns = self::resolveColumns($indexedRow, $options, $headerLookup);
                $columnsInitialized = true;
            }

            // Check empty row
            if (RowMapper::isRowEmpty($indexedRow, $resolvedColumns)) {
                if ($options->stopAtEmptyRow) {
                    break;
                }
                if ($options->skipEmptyRows) {
                    continue;
                }
            }

            $mapped = RowMapper::map($lineNum, $indexedRow, $resolvedColumns, $options, $globalOptions);

            if ($mapped['errors'] !== []) {
                if ($options->onError === 'throw') {
                    throw RowValidationException::fromRowError($mapped['errors'][0]);
                }

                foreach ($mapped['errors'] as $err) {
                    $currentChunkErrors[] = $err;
                    $allErrors[] = $err;
                }

                if ($options->onError === 'skip_row') {
                    continue;
                }
            }

            $currentChunk[] = $mapped['row'];
            $totalRowsProcessed++;

            if ($options->maxRows !== null && $totalRowsProcessed > $options->maxRows) {
                throw ReaderException::maxRowsExceeded($options->maxRows, $totalRowsProcessed);
            }

            if (count($currentChunk) >= $chunkSize) {
                $chunkContext = new ChunkContext(
                    chunkIndex: $chunkIndex,
                    startRow: $lineNum - count($currentChunk) + 1,
                    endRow: $lineNum,
                    rowCount: count($currentChunk),
                    totalRowsProcessed: $totalRowsProcessed,
                    errors: $currentChunkErrors,
                );

                $ret = $callback($currentChunk, $chunkContext);
                $chunkIndex++;
                $currentChunk = [];
                $currentChunkErrors = [];

                if ($ret === false || $chunkContext->shouldStop()) {
                    $stopStreaming = true;
                    break;
                }
            }
        }

        // Emit final remainder chunk if any
        if (!$stopStreaming && $currentChunk !== []) {
            $chunkContext = new ChunkContext(
                chunkIndex: $chunkIndex,
                startRow: $lineNum - count($currentChunk) + 1,
                endRow: $lineNum,
                rowCount: count($currentChunk),
                totalRowsProcessed: $totalRowsProcessed,
                errors: $currentChunkErrors,
            );

            $callback($currentChunk, $chunkContext);
            $chunkIndex++;
        }

        return new ReadSummary(
            totalRows: $totalRowsProcessed,
            totalChunks: $chunkIndex,
            errors: $allErrors,
            headers: $headers,
            sheetName: 'Worksheet',
        );
    }

    /**
     * @param array<string, string> $headers
     * @return array<string, int>
     */
    private static function buildHeaderLookup(array $headers): array
    {
        $lookup = [];
        foreach ($headers as $letter => $val) {
            $colIdx = Coordinate::columnIndex($letter);
            $norm = mb_strtolower(trim($val));
            if ($norm !== '' && !isset($lookup[$norm])) {
                $lookup[$norm] = $colIdx;
            }
        }

        return $lookup;
    }

    /**
     * @param array<int, mixed> $rawRow
     * @param array<string, int>|null $headerLookup
     * @return list<ReadColumn>
     */
    private static function resolveColumns(
        array $rawRow,
        ReadOptions $options,
        ?array $headerLookup,
    ): array {
        if ($options->columns === []) {
            $resolved = [];
            foreach (array_keys($rawRow) as $cIdx) {
                $resolved[] = new ReadColumn(
                    key: Coordinate::columnLetter($cIdx),
                    column: $cIdx,
                    type: 'raw',
                );
            }

            return $resolved;
        }

        $resolved = [];
        foreach ($options->columns as $col) {
            if ($col->column !== null) {
                $colIndex = Coordinate::columnIndex($col->column);
                $resolved[] = $col->withResolvedColumn($colIndex);
                continue;
            }

            if ($col->header !== null) {
                $headersToMatch = is_array($col->header) ? $col->header : [$col->header];
                $matchedColIndex = null;

                if ($headerLookup !== null) {
                    foreach ($headersToMatch as $headerText) {
                        $norm = mb_strtolower(trim($headerText));
                        if (isset($headerLookup[$norm])) {
                            $matchedColIndex = $headerLookup[$norm];
                            break;
                        }
                    }
                }

                if ($matchedColIndex !== null) {
                    $resolved[] = $col->withResolvedColumn($matchedColIndex);
                } elseif (!$col->optional) {
                    $displayHeader = is_array($col->header) ? implode(' / ', $col->header) : $col->header;

                    throw ReaderException::missingHeaderColumn($displayHeader);
                } else {
                    $resolved[] = $col->withResolvedColumn(null);
                }
            }
        }

        return $resolved;
    }

    /**
     * Resolve CSV delimiter, auto-detecting from candidate list if 'auto'.
     */
    public static function resolveDelimiter(string $path, string $declared): string
    {
        if ($declared !== 'auto') {
            return $declared;
        }

        $handle = fopen($path, 'r');
        if ($handle === false) {
            return ',';
        }

        $lines = [];
        for ($i = 0; $i < 5 && !feof($handle); $i++) {
            $line = fgets($handle);
            if ($line !== false && trim($line) !== '') {
                $lines[] = $line;
            }
        }
        fclose($handle);

        if ($lines === []) {
            return ',';
        }

        $bestDelimiter = ',';
        $maxCount = 0;

        foreach (self::DELIMITER_CANDIDATES as $candidate) {
            $total = 0;
            foreach ($lines as $line) {
                $total += substr_count($line, $candidate);
            }
            if ($total > $maxCount) {
                $maxCount = $total;
                $bestDelimiter = $candidate;
            }
        }

        return $bestDelimiter;
    }

    /**
     * Resolve input encoding (BOM or auto-detection).
     */
    public static function resolveEncoding(string $path, string $declared): string
    {
        if ($declared !== 'auto') {
            return strtoupper($declared);
        }

        $handle = fopen($path, 'r');
        if ($handle === false) {
            return 'UTF-8';
        }

        $bom = fread($handle, 4);
        fclose($handle);

        if ($bom !== false) {
            if (str_starts_with($bom, "\xEF\xBB\xBF")) {
                return 'UTF-8';
            }
            if (str_starts_with($bom, "\xFF\xFE")) {
                return 'UTF-16LE';
            }
            if (str_starts_with($bom, "\xFE\xFF")) {
                return 'UTF-16BE';
            }
        }

        // Sample first 2048 bytes
        $sample = file_get_contents($path, false, null, 0, 2048);
        if ($sample === false) {
            return 'UTF-8';
        }

        $detected = mb_detect_encoding($sample, ['UTF-8', 'Windows-1252', 'ISO-8859-1'], true);

        return $detected !== false ? $detected : 'UTF-8';
    }

    /**
     * @param list<mixed> $row
     * @return list<mixed>
     */
    private static function stripBomFromRow(array $row): array
    {
        if (isset($row[0]) && is_string($row[0])) {
            $row[0] = preg_replace('/^\xEF\xBB\xBF/', '', $row[0]) ?? $row[0];
        }

        return $row;
    }

}
