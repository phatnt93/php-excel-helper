<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

/**
 * Sanitizes worksheet names according to Excel's 31-character limit, forbidden characters, and deduplication.
 */
final class SheetNameSanitizer
{
    private const FORBIDDEN_CHARS = ['*', ':', '?', '/', '\\', '[', ']'];
    private const MAX_LENGTH = 31;

    /**
     * @param list<string> $existingNames Already used sheet names
     */
    public static function sanitize(string $name, array $existingNames = []): string
    {
        $clean = str_replace(self::FORBIDDEN_CHARS, '_', trim($name));
        $clean = trim($clean, "'");

        if ($clean === '') {
            $clean = 'Sheet';
        }

        if (mb_strlen($clean) > self::MAX_LENGTH) {
            $clean = mb_substr($clean, 0, self::MAX_LENGTH);
        }

        // Check deduplication (Excel sheet names are case-insensitive)
        $lowerExisting = array_map('mb_strtolower', $existingNames);
        if (!in_array(mb_strtolower($clean), $lowerExisting, true)) {
            return $clean;
        }

        $counter = 2;
        while (true) {
            $suffix = ' (' . $counter . ')';
            $suffixLen = mb_strlen($suffix);
            $baseMaxLen = self::MAX_LENGTH - $suffixLen;

            $candidate = (mb_strlen($clean) > $baseMaxLen ? mb_substr($clean, 0, $baseMaxLen) : $clean) . $suffix;

            if (!in_array(mb_strtolower($candidate), $lowerExisting, true)) {
                return $candidate;
            }

            $counter++;
        }
    }
}
