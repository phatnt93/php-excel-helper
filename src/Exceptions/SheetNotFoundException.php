<?php

declare(strict_types=1);

namespace ExcelHelper\Exceptions;

use InvalidArgumentException;

/**
 * Thrown when a requested worksheet is not found by index or by name.
 */
class SheetNotFoundException extends InvalidArgumentException implements ExcelHelperException
{
    /**
     * @param list<string> $available
     */
    public static function byIndex(int $index, array $available): self
    {
        $availableList = implode(', ', array_map(static fn(string $name, int $i): string => "#{$i}: '{$name}'", $available, array_keys($available)));

        return new self(sprintf(
            'Worksheet at index %d was not found. Available sheets: %s.',
            $index,
            $availableList !== '' ? $availableList : '(none)',
        ));
    }

    /**
     * @param list<string> $available
     */
    public static function byName(string $name, array $available): self
    {
        $availableList = implode(', ', array_map(static fn(string $n): string => "'{$n}'", $available));

        return new self(sprintf(
            'Worksheet with name "%s" was not found. Available sheets: %s.',
            $name,
            $availableList !== '' ? $availableList : '(none)',
        ));
    }
}
