<?php

declare(strict_types=1);

namespace ExcelHelper\Exceptions;

use RuntimeException;
use Throwable;

/**
 * Thrown when writing a spreadsheet fails (directory not writable, headers sent, limit exceeded, save failure).
 */
class WriterException extends RuntimeException implements ExcelHelperException
{
    public static function directoryNotWritable(string $dir): self
    {
        return new self(sprintf('Target directory is not writable: "%s".', $dir));
    }

    public static function headersAlreadySent(): self
    {
        return new self('Cannot send spreadsheet download because HTTP headers have already been sent.');
    }

    public static function unsupportedTypeForWriting(string $type, string $reason): self
    {
        return new self(sprintf('Type "%s" cannot be used for writing: %s', $type, $reason));
    }

    public static function xlsLimitExceeded(int $rows, int $cols, int $maxRows = 65536, int $maxCols = 256): self
    {
        return new self(sprintf(
            'XLS (BIFF8) format limit exceeded: got %d rows and %d columns. Maximum allowed is %d rows and %d columns.',
            $rows,
            $cols,
            $maxRows,
            $maxCols,
        ));
    }

    public static function failedToSave(string $path, ?Throwable $previous = null): self
    {
        return new self(sprintf('Failed to save spreadsheet to "%s".', $path), 0, $previous);
    }

    public static function cannotWriteValue(string $cell, string $reason): self
    {
        return new self(sprintf('Cannot write value to cell "%s": %s', $cell, $reason));
    }

    public static function withMessage(string $message, ?Throwable $previous = null): self
    {
        return new self($message, 0, $previous);
    }
}
