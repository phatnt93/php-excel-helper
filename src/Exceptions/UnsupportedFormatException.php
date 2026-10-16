<?php

declare(strict_types=1);

namespace ExcelHelper\Exceptions;

use InvalidArgumentException;

/**
 * Thrown when an unsupported spreadsheet format is encountered or requested.
 */
class UnsupportedFormatException extends InvalidArgumentException implements ExcelHelperException
{
    /**
     * @param list<string> $allowed
     */
    public static function notWhitelisted(string $format, array $allowed): self
    {
        return new self(sprintf(
            'Format "%s" is not in the allowed formats whitelist (%s).',
            $format,
            implode(', ', $allowed),
        ));
    }

    public static function unrecognized(string $path): self
    {
        return new self(sprintf('Could not recognize the spreadsheet format for file "%s".', $path));
    }

    public static function unsupportedForReading(string $format): self
    {
        return new self(sprintf('Format "%s" is not supported for reading.', $format));
    }

    public static function unsupportedForWriting(string $format): self
    {
        return new self(sprintf('Format "%s" is not supported for writing.', $format));
    }
}
