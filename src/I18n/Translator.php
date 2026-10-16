<?php

declare(strict_types=1);

namespace ExcelHelper\I18n;

use ExcelHelper\Exceptions\InvalidConfigException;

/**
 * Handles translation of row validation error messages and table labels.
 */
final class Translator
{
    public const DEFAULT_LOCALE = 'en';

    /**
     * @var list<string>
     */
    public const SUPPORTED_LOCALES = ['en', 'vi'];

    private static string $currentLocale = self::DEFAULT_LOCALE;

    /**
     * @var array<string, array<string, string>>
     */
    private static array $catalogs = [];

    /**
     * Set the current global locale.
     *
     * @throws InvalidConfigException
     */
    public static function setLocale(string $locale): void
    {
        $normalized = strtolower(trim($locale));

        if (!self::isLocaleSupported($normalized)) {
            throw InvalidConfigException::withMessage(sprintf(
                'Unsupported locale "%s". Supported locales are: %s.',
                $locale,
                implode(', ', self::SUPPORTED_LOCALES),
            ));
        }

        self::$currentLocale = $normalized;
    }

    public static function getLocale(): string
    {
        return self::$currentLocale;
    }

    public static function reset(): void
    {
        self::$currentLocale = self::DEFAULT_LOCALE;
        self::$catalogs = [];
    }

    public static function isLocaleSupported(string $locale): bool
    {
        return in_array(strtolower(trim($locale)), self::SUPPORTED_LOCALES, true);
    }

    /**
     * @return list<string>
     */
    public static function availableLocales(): array
    {
        return self::SUPPORTED_LOCALES;
    }

    /**
     * Translate a message key into the specified (or current) locale with parameter replacement.
     *
     * @param array<string, mixed> $params
     *
     * @throws InvalidConfigException if an explicit unsupported locale is requested
     */
    public static function trans(string $key, array $params = [], ?string $locale = null): string
    {
        $targetLocale = $locale !== null ? strtolower(trim($locale)) : self::$currentLocale;

        if (!self::isLocaleSupported($targetLocale)) {
            throw InvalidConfigException::withMessage(sprintf(
                'Unsupported locale "%s". Supported locales are: %s.',
                $locale,
                implode(', ', self::SUPPORTED_LOCALES),
            ));
        }

        $catalog = self::loadCatalog($targetLocale);
        $template = $catalog[$key] ?? null;

        if ($template === null && $targetLocale !== self::DEFAULT_LOCALE) {
            $fallback = self::loadCatalog(self::DEFAULT_LOCALE);
            $template = $fallback[$key] ?? null;
        }

        if ($template === null) {
            $template = $key;
        }

        return self::interpolate($template, $params);
    }

    /**
     * @return array<string, string>
     */
    public static function loadCatalog(string $locale): array
    {
        if (isset(self::$catalogs[$locale])) {
            return self::$catalogs[$locale];
        }

        $file = __DIR__ . '/' . $locale . '.php';
        if (is_file($file)) {
            /** @var array<string, string> $loaded */
            $loaded = require $file;
            self::$catalogs[$locale] = $loaded;

            return $loaded;
        }

        return [];
    }

    /**
     * Interpolate parameters like {row}, {column}, {value} into a template string.
     *
     * @param array<string, mixed> $params
     */
    private static function interpolate(string $template, array $params): string
    {
        if ($params === []) {
            return $template;
        }

        $replace = [];
        foreach ($params as $k => $v) {
            if ($v === null) {
                $str = '';
            } elseif (is_bool($v)) {
                $str = $v ? 'true' : 'false';
            } elseif (is_scalar($v) || (is_object($v) && method_exists($v, '__toString'))) {
                $str = (string) $v;
            } else {
                $json = json_encode($v, JSON_UNESCAPED_UNICODE);
                $str = $json !== false ? $json : '';
            }
            $replace['{' . $k . '}'] = $str;
        }

        return strtr($template, $replace);
    }
}
