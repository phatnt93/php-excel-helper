<?php

declare(strict_types=1);

namespace ExcelHelper\Config;

use ExcelHelper\Exceptions\InvalidConfigException;

/**
 * Immutable value object representing a dot-separated configuration path.
 */
final class ConfigPath
{
    /**
     * @param list<string|int> $segments
     */
    public function __construct(
        private readonly array $segments = [],
    ) {}

    public static function root(string $root = ''): self
    {
        return $root === '' ? new self([]) : new self([$root]);
    }

    public function append(string|int $segment): self
    {
        $copy = $this->segments;
        $copy[] = $segment;

        return new self($copy);
    }

    /**
     * @return list<string|int>
     */
    public function segments(): array
    {
        return $this->segments;
    }

    public function isEmpty(): bool
    {
        return $this->segments === [];
    }

    public function toString(): string
    {
        return implode('.', array_map('strval', $this->segments));
    }

    public function __toString(): string
    {
        return $this->toString();
    }

    /**
     * Validate that all array keys exist within the allowed keys list.
     *
     * @param array<array-key, mixed> $config
     * @param list<string> $allowed
     *
     * @throws InvalidConfigException
     */
    public static function assertKnownKeys(array $config, array $allowed, self $path): void
    {
        $allowedLookup = array_flip($allowed);

        foreach (array_keys($config) as $key) {
            $keyStr = (string) $key;
            if (!isset($allowedLookup[$keyStr])) {
                $subPath = $path->append($keyStr);
                $allowedList = implode(', ', array_map(static fn(string $k): string => "'{$k}'", $allowed));

                throw InvalidConfigException::atPath(
                    $subPath,
                    sprintf('Unknown configuration option "%s". Allowed options are: %s.', $keyStr, $allowedList),
                );
            }
        }
    }
}
