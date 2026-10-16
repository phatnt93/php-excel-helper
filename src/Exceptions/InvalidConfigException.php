<?php

declare(strict_types=1);

namespace ExcelHelper\Exceptions;

use ExcelHelper\Config\ConfigPath;
use InvalidArgumentException;

/**
 * Thrown when an invalid configuration is supplied (e.g. unknown keys, invalid types, conflicting options).
 */
class InvalidConfigException extends InvalidArgumentException implements ExcelHelperException
{
    private ?ConfigPath $configPath = null;

    public static function atPath(ConfigPath $path, string $reason): self
    {
        $pathString = $path->toString();
        $prefix = $pathString !== '' ? sprintf('Configuration error at "%s": ', $pathString) : 'Configuration error: ';

        $exception = new self($prefix . $reason);
        $exception->configPath = $path;

        return $exception;
    }

    public static function withMessage(string $message): self
    {
        return new self($message);
    }

    public function getConfigPath(): ?ConfigPath
    {
        return $this->configPath;
    }
}
