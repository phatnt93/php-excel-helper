<?php

declare(strict_types=1);

namespace ExcelHelper\Support;

use ExcelHelper\Exceptions\InvalidConfigException;

/**
 * Normalizes color hex strings into PhpSpreadsheet-compatible RRGGBB uppercase format.
 */
final class Color
{
    /**
     * Normalize #RGB, #RRGGBB, RGB, or RRGGBB hex into 6-character uppercase RRGGBB.
     *
     * @throws InvalidConfigException
     */
    public static function normalize(string $color): string
    {
        $trimmed = ltrim(trim($color), '#');

        if (preg_match('/^[0-9a-fA-F]{3}$/D', $trimmed)) {
            $r = $trimmed[0];
            $g = $trimmed[1];
            $b = $trimmed[2];

            return strtoupper($r . $r . $g . $g . $b . $b);
        }

        if (preg_match('/^[0-9a-fA-F]{6}$/D', $trimmed)) {
            return strtoupper($trimmed);
        }

        throw InvalidConfigException::withMessage(sprintf(
            'Invalid hex color format "%s". Expected format "#RGB", "#RRGGBB", "RGB" or "RRGGBB".',
            $color,
        ));
    }
}
