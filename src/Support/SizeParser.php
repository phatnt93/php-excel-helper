<?php

declare(strict_types=1);

namespace ExcelHelper\Support;

use ExcelHelper\Exceptions\InvalidConfigException;

/**
 * Parses byte sizes from integers or strings (e.g. '20M', '512KB', '1G').
 */
final class SizeParser
{
    private const UNITS = [
        'B' => 1,
        'K' => 1024,
        'KB' => 1024,
        'M' => 1048576,
        'MB' => 1048576,
        'G' => 1073741824,
        'GB' => 1073741824,
    ];

    /**
     * Parse human-readable or integer size into bytes.
     *
     * @throws InvalidConfigException
     */
    public static function parse(int|string|null $size): ?int
    {
        if ($size === null) {
            return null;
        }

        if (is_int($size)) {
            if ($size < 0) {
                throw InvalidConfigException::withMessage(sprintf('Size cannot be negative, got %d.', $size));
            }

            return $size;
        }

        $trimmed = trim($size);

        if ($trimmed === '') {
            return null;
        }

        if (!preg_match('/^(\d+(?:\.\d+)?)\s*([A-Za-z]+)?$/D', $trimmed, $matches)) {
            throw InvalidConfigException::withMessage(sprintf('Invalid size format: "%s". Expected e.g. "20M", "512KB", "1048576".', $size));
        }

        $value = (float) $matches[1];
        $unit = isset($matches[2]) ? strtoupper($matches[2]) : 'B';

        if (!isset(self::UNITS[$unit])) {
            throw InvalidConfigException::withMessage(sprintf('Unknown size unit "%s" in "%s". Supported units: B, K, KB, M, MB, G, GB.', $unit, $size));
        }

        return (int) round($value * self::UNITS[$unit]);
    }
}
