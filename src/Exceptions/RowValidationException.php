<?php

declare(strict_types=1);

namespace ExcelHelper\Exceptions;

use ExcelHelper\Reader\RowError;
use RuntimeException;

/**
 * Thrown when row validation fails under 'throw' on_error policy.
 */
class RowValidationException extends RuntimeException implements ExcelHelperException
{
    private RowError $error;

    public static function fromRowError(RowError $error): self
    {
        $message = sprintf(
            'Row %d column "%s" (%s) validation failed [%s]: %s',
            $error->row,
            $error->column,
            $error->key,
            $error->code,
            $error->message,
        );

        $exception = new self($message);
        $exception->error = $error;

        return $exception;
    }

    public function getRowError(): RowError
    {
        return $this->error;
    }
}
