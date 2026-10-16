<?php

declare(strict_types=1);

namespace ExcelHelper\Support;

use ExcelHelper\Exceptions\InvalidConfigException;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate as PhpSpreadsheetCoordinate;
use Throwable;

/**
 * Coordinate helpers wrapping PhpSpreadsheet's Coordinate with strict bounds checking.
 */
final class Coordinate
{
    public const MAX_COLUMN_INDEX = 16384; // XFD
    public const MAX_COLUMN_LETTER = 'XFD';

    /**
     * Determine if a string is a valid Excel column letter (A..XFD). Case-insensitive.
     */
    public static function isColumnLetter(string $column): bool
    {
        $upper = strtoupper(trim($column));

        if (!preg_match('/^[A-Z]{1,3}$/D', $upper)) {
            return false;
        }

        try {
            $index = PhpSpreadsheetCoordinate::columnIndexFromString($upper);

            return $index >= 1 && $index <= self::MAX_COLUMN_INDEX;
        } catch (Throwable) {
            return false;
        }
    }

    /**
     * Convert a column letter or 1-based column number to a 1-based column index.
     *
     * @throws InvalidConfigException
     */
    public static function columnIndex(string|int $column): int
    {
        if (is_int($column)) {
            if ($column < 1 || $column > self::MAX_COLUMN_INDEX) {
                throw InvalidConfigException::withMessage(sprintf(
                    'Column index must be between 1 and %d (%s), got %d.',
                    self::MAX_COLUMN_INDEX,
                    self::MAX_COLUMN_LETTER,
                    $column,
                ));
            }

            return $column;
        }

        $trimmed = trim($column);

        if (is_numeric($trimmed) && (string) (int) $trimmed === $trimmed) {
            return self::columnIndex((int) $trimmed);
        }

        $upper = strtoupper($trimmed);

        if (!self::isColumnLetter($upper)) {
            throw InvalidConfigException::withMessage(sprintf(
                'Invalid column identifier "%s". Must be a 1-based integer or a valid column letter (A..%s).',
                $column,
                self::MAX_COLUMN_LETTER,
            ));
        }

        return PhpSpreadsheetCoordinate::columnIndexFromString($upper);
    }

    /**
     * Convert a 1-based column index or column letter to uppercase column letter.
     *
     * @throws InvalidConfigException
     */
    public static function columnLetter(int|string $index): string
    {
        if (is_string($index)) {
            $trimmed = trim($index);
            if (is_numeric($trimmed) && (string) (int) $trimmed === $trimmed) {
                return self::columnLetter((int) $trimmed);
            }

            $upper = strtoupper($trimmed);
            if (self::isColumnLetter($upper)) {
                return $upper;
            }

            throw InvalidConfigException::withMessage(sprintf(
                'Invalid column identifier "%s". Must be a 1-based integer or a valid column letter (A..%s).',
                $index,
                self::MAX_COLUMN_LETTER,
            ));
        }

        if ($index < 1 || $index > self::MAX_COLUMN_INDEX) {
            throw InvalidConfigException::withMessage(sprintf(
                'Column index must be between 1 and %d (%s), got %d.',
                self::MAX_COLUMN_INDEX,
                self::MAX_COLUMN_LETTER,
                $index,
            ));
        }

        return PhpSpreadsheetCoordinate::stringFromColumnIndex($index);
    }

    /**
     * Parse a cell address (e.g. 'B3', '$C$5') into 1-based column and row.
     *
     * @return array{col: int, row: int}
     *
     * @throws InvalidConfigException
     */
    public static function parseCell(string $cell): array
    {
        $clean = str_replace('$', '', trim($cell));

        try {
            /** @var array{0: int, 1: int} $indexes */
            $indexes = PhpSpreadsheetCoordinate::indexesFromString($clean);

            $col = $indexes[0];
            $row = $indexes[1];

            if ($col < 1 || $col > self::MAX_COLUMN_INDEX || $row < 1) {
                throw InvalidConfigException::withMessage(sprintf('Cell coordinates out of bounds: "%s".', $cell));
            }

            return ['col' => $col, 'row' => $row];
        } catch (Throwable $e) {
            if ($e instanceof InvalidConfigException) {
                throw $e;
            }

            throw InvalidConfigException::withMessage(sprintf('Invalid cell address: "%s".', $cell));
        }
    }

    /**
     * Format 1-based column and row into a cell address (e.g. 2, 3 -> 'B3').
     *
     * @throws InvalidConfigException
     */
    public static function formatCell(int $col, int $row): string
    {
        if ($row < 1) {
            throw InvalidConfigException::withMessage(sprintf('Row index must be >= 1, got %d.', $row));
        }

        return self::columnLetter($col) . $row;
    }
}
