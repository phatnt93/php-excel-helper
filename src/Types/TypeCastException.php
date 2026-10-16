<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Exceptions\ExcelHelperException;
use RuntimeException;
use Throwable;

/**
 * Thrown internally by TypeInterface implementations when conversion of a cell value fails.
 */
class TypeCastException extends RuntimeException implements ExcelHelperException
{
    /**
     * @param array<string, mixed> $params
     */
    public function __construct(
        private readonly string $codeString,
        private readonly array $params = [],
        string $message = '',
        ?Throwable $previous = null,
    ) {
        parent::__construct($message !== '' ? $message : sprintf('Type cast failed with code "%s".', $codeString), 0, $previous);
    }

    public function getCodeString(): string
    {
        return $this->codeString;
    }

    /**
     * @return array<string, mixed>
     */
    public function getParams(): array
    {
        return $this->params;
    }
}
