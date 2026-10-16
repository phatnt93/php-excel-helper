<?php

declare(strict_types=1);

namespace ExcelHelper\Exceptions;

use RuntimeException;
use Throwable;

/**
 * Thrown when reading a spreadsheet fails (file not found, unreadable, corrupt, limits exceeded, missing header).
 */
class ReaderException extends RuntimeException implements ExcelHelperException
{
    public static function fileNotFound(string $path): self
    {
        return new self(sprintf('Spreadsheet file not found: "%s".', $path));
    }

    public static function notReadable(string $path): self
    {
        return new self(sprintf('Spreadsheet file is not readable: "%s".', $path));
    }

    public static function isADirectory(string $path): self
    {
        return new self(sprintf('Path is a directory, not a file: "%s".', $path));
    }

    public static function maxFileSizeExceeded(string $path, int $limit, int $actual): self
    {
        return new self(sprintf(
            'Spreadsheet file "%s" exceeds maximum allowed size of %d bytes (actual: %d bytes).',
            $path,
            $limit,
            $actual,
        ));
    }

    public static function maxRowsExceeded(int $limit, int $actual): self
    {
        return new self(sprintf(
            'Spreadsheet exceeds maximum allowed rows of %d (actual: %d rows).',
            $limit,
            $actual,
        ));
    }

    public static function missingHeaderColumn(string $header): self
    {
        return new self(sprintf('Required header column "%s" was not found in the header row.', $header));
    }

    public static function cannotOpenFile(string $path, ?Throwable $previous = null): self
    {
        return new self(sprintf('Cannot open spreadsheet file "%s".', $path), 0, $previous);
    }

    public static function formulaError(string $cell, string $message, ?Throwable $previous = null): self
    {
        return new self(sprintf('Error calculating formula in cell "%s": %s', $cell, $message), 0, $previous);
    }

    public static function emptyFile(string $path): self
    {
        return new self(sprintf('Spreadsheet file is empty: "%s".', $path));
    }

    public static function corruptedFile(string $path, ?Throwable $previous = null): self
    {
        return new self(sprintf('Spreadsheet file is corrupted or unreadable: "%s".', $path), 0, $previous);
    }

    public static function withMessage(string $message, ?Throwable $previous = null): self
    {
        return new self($message, 0, $previous);
    }
}
