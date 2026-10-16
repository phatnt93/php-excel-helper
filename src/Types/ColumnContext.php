<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

/**
 * Contextual configuration passed to TypeInterface::read() and TypeInterface::write().
 */
final class ColumnContext
{
    /**
     * @param array<string, mixed> $options Column-specific options
     * @param array<string, mixed> $globalOptions Global reader/writer options
     */
    public function __construct(
        public readonly string $key,
        public readonly array $options = [],
        public readonly array $globalOptions = [],
    ) {}

    /**
     * Get an option value, looking up column options first, then global options, then default.
     */
    public function get(string $name, mixed $default = null): mixed
    {
        if (array_key_exists($name, $this->options)) {
            return $this->options[$name];
        }

        if (array_key_exists($name, $this->globalOptions)) {
            return $this->globalOptions[$name];
        }

        return $default;
    }

    public function isTrim(): bool
    {
        return (bool) $this->get('trim', true);
    }

    public function getDateFormat(): ?string
    {
        if (array_key_exists('format', $this->options)) {
            /** @var ?string $format */
            $format = $this->options['format'];

            return $format;
        }

        if (array_key_exists('date_format', $this->globalOptions)) {
            /** @var ?string $format */
            $format = $this->globalOptions['date_format'];

            return $format;
        }

        return 'Y-m-d';
    }

    public function getDateTimeFormat(): ?string
    {
        if (array_key_exists('format', $this->options)) {
            /** @var ?string $format */
            $format = $this->options['format'];

            return $format;
        }

        if (array_key_exists('datetime_format', $this->globalOptions)) {
            /** @var ?string $format */
            $format = $this->globalOptions['datetime_format'];

            return $format;
        }

        return 'Y-m-d H:i:s';
    }

    public function isDateSystem1904(): bool
    {
        return (bool) ($this->globalOptions['date_system_1904'] ?? false);
    }

    public function getLocale(): string
    {
        return (string) ($this->globalOptions['locale'] ?? 'en');
    }
}
