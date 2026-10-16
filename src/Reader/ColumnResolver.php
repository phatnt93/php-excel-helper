<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Support\Coordinate;
use Normalizer;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Resolves column specifications (letters, numbers, sequential, header text, auto-detect)
 * into concrete 1-based column indices.
 */
final class ColumnResolver
{
    /**
     * @return list<ReadColumn>
     *
     * @throws ReaderException
     */
    public static function resolve(Worksheet $worksheet, ReadOptions $options): array
    {
        if ($options->columns === []) {
            return self::autoDetectColumns($worksheet, $options);
        }

        $resolved = [];
        $headerLookup = null;

        foreach ($options->columns as $col) {
            if ($col->column !== null) {
                $colIndex = Coordinate::columnIndex($col->column);
                $resolved[] = $col->withResolvedColumn($colIndex);
                continue;
            }

            if ($col->header !== null) {
                if ($headerLookup === null) {
                    $headerLookup = self::buildHeaderLookup($worksheet, $options->headerRow ?? 1);
                }

                $headersToMatch = is_array($col->header) ? $col->header : [$col->header];
                $matchedColIndex = null;

                foreach ($headersToMatch as $headerText) {
                    $norm = self::normalizeHeaderText($headerText);
                    if (isset($headerLookup[$norm])) {
                        $matchedColIndex = $headerLookup[$norm];
                        break;
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
     * @return list<ReadColumn>
     */
    private static function autoDetectColumns(Worksheet $worksheet, ReadOptions $options): array
    {
        $highestCol = $worksheet->getHighestDataColumn();
        $maxIndex = Coordinate::columnIndex($highestCol);

        $resolved = [];
        $seenKeys = [];

        for ($c = 1; $c <= $maxIndex; $c++) {
            $colLetter = Coordinate::columnLetter($c);

            if ($options->headerRow !== null) {
                $rawVal = $worksheet->getCell($colLetter . $options->headerRow)->getValue();
                $headerStr = trim((string) $rawVal);

                if ($headerStr === '') {
                    $key = 'column_' . $colLetter;
                } elseif (isset($seenKeys[$headerStr])) {
                    $key = $headerStr . '_' . $colLetter;
                } else {
                    $key = $headerStr;
                }
            } else {
                $key = $colLetter;
            }

            $seenKeys[$key] = true;
            $resolved[] = new ReadColumn(
                key: $key,
                column: $c,
                type: 'raw',
            );
        }

        return $resolved;
    }

    /**
     * @return array<string, int> normalized header -> 1-based column index
     */
    private static function buildHeaderLookup(Worksheet $worksheet, int $headerRow): array
    {
        $highestCol = $worksheet->getHighestDataColumn();
        $maxIndex = Coordinate::columnIndex($highestCol);

        $lookup = [];
        for ($c = 1; $c <= $maxIndex; $c++) {
            $colLetter = Coordinate::columnLetter($c);
            $rawVal = $worksheet->getCell($colLetter . $headerRow)->getValue();
            if ($rawVal !== null && $rawVal !== '') {
                $norm = self::normalizeHeaderText((string) $rawVal);
                if (!isset($lookup[$norm])) {
                    $lookup[$norm] = $c;
                }
            }
        }

        return $lookup;
    }

    private static function normalizeHeaderText(string $text): string
    {
        $trimmed = trim($text);

        if (class_exists(Normalizer::class)) {
            $normalized = Normalizer::normalize($trimmed, Normalizer::FORM_C);
            if ($normalized !== false) {
                $trimmed = $normalized;
            }
        }

        return mb_strtolower($trimmed);
    }
}
