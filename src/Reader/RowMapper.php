<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use ExcelHelper\I18n\Translator;
use ExcelHelper\Support\Coordinate;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\TypeCastException;
use ExcelHelper\Types\TypeRegistry;

/**
 * Maps a single raw spreadsheet row into the final associative array according to column rules.
 */
final class RowMapper
{
    /**
     * @param int $rowNumber 1-based Excel row number
     * @param array<int, mixed> $rawRow Raw cell values indexed by 1-based column index
     * @param list<ReadColumn> $columns
     * @param array<string, mixed> $globalOptions
     * @return array{row: array<string, mixed>, errors: list<RowError>}
     */
    public static function map(
        int $rowNumber,
        array $rawRow,
        array $columns,
        ReadOptions $options,
        array $globalOptions,
    ): array {
        $mapped = [];
        $errors = [];
        $locale = $options->locale;

        foreach ($columns as $col) {
            if ($col->column === null) {
                $mapped[$col->key] = $col->default;
                continue;
            }

            $colIndex = (int) $col->column;
            $colLetter = Coordinate::columnLetter($colIndex);
            $rawVal = $rawRow[$colIndex] ?? null;

            // 1. Trim string values
            $shouldTrim = $col->trim ?? $options->trim;
            if ($shouldTrim && is_string($rawVal)) {
                $rawVal = trim($rawVal);
            }

            // 2. Empty check
            $isEmpty = $rawVal === null || $rawVal === '';

            if ($isEmpty) {
                if ($col->required) {
                    $msg = Translator::trans('row.required', [
                        'row' => $rowNumber,
                        'column' => $colLetter,
                        'key' => $col->key,
                    ], $locale);

                    $errors[] = new RowError(
                        row: $rowNumber,
                        column: $colLetter,
                        key: $col->key,
                        code: 'required',
                        value: $rawVal,
                        message: $msg,
                    );
                }

                $mapped[$col->key] = $col->default;
                continue;
            }

            // 3. Dictionary mapping
            if ($col->map !== null) {
                $matched = false;
                $mappedVal = null;

                if (is_scalar($rawVal) && isset($col->map[(string) $rawVal])) {
                    $mappedVal = $col->map[(string) $rawVal];
                    $matched = true;
                } else {
                    $rawValLower = is_string($rawVal) ? mb_strtolower($rawVal) : null;
                    foreach ($col->map as $mapKey => $mapTarget) {
                        if ($rawValLower !== null && mb_strtolower((string) $mapKey) === $rawValLower) {
                            $mappedVal = $mapTarget;
                            $matched = true;
                            break;
                        }
                    }
                }

                if ($matched) {
                    $rawVal = $mappedVal;
                } elseif ($col->mapStrict) {
                    $msg = Translator::trans('row.not_in_map', [
                        'row' => $rowNumber,
                        'column' => $colLetter,
                        'key' => $col->key,
                        'value' => $rawVal,
                    ], $locale);

                    $errors[] = new RowError(
                        row: $rowNumber,
                        column: $colLetter,
                        key: $col->key,
                        code: 'not_in_map',
                        value: $rawVal,
                        message: $msg,
                    );

                    $mapped[$col->key] = $col->default;
                    continue;
                }
            }

            // 4. Type cast
            $type = TypeRegistry::get($col->type);
            $colCtx = new ColumnContext($col->key, $col->options, $globalOptions);

            try {
                $castedVal = $type->read($rawVal, $colCtx);
            } catch (TypeCastException $e) {
                $transKey = 'row.' . $e->getCodeString();
                $params = array_merge([
                    'row' => $rowNumber,
                    'column' => $colLetter,
                    'key' => $col->key,
                    'value' => $rawVal,
                ], $e->getParams());

                $msg = Translator::trans($transKey, $params, $locale);

                $errors[] = new RowError(
                    row: $rowNumber,
                    column: $colLetter,
                    key: $col->key,
                    code: $e->getCodeString(),
                    value: $rawVal,
                    message: $msg,
                );

                $mapped[$col->key] = $col->default;
                continue;
            }

            // 5. Transform callback
            if ($col->transform !== null) {
                $rowCtx = new RowContext($rowNumber, $rawRow, $mapped);
                $castedVal = ($col->transform)($castedVal, $rowCtx);
            }

            // 6. Validate callback
            if ($col->validate !== null) {
                $rowCtx = new RowContext($rowNumber, $rawRow, $mapped);
                $validResult = ($col->validate)($castedVal, $rowCtx);

                if ($validResult !== true) {
                    $customMsg = is_string($validResult) && $validResult !== ''
                        ? $validResult
                        : 'Validation failed';

                    $msg = Translator::trans('row.custom', [
                        'row' => $rowNumber,
                        'column' => $colLetter,
                        'key' => $col->key,
                        'message' => $customMsg,
                    ], $locale);

                    $errors[] = new RowError(
                        row: $rowNumber,
                        column: $colLetter,
                        key: $col->key,
                        code: 'custom',
                        value: $rawVal,
                        message: $msg,
                    );

                    $mapped[$col->key] = $col->default;
                    continue;
                }
            }

            $mapped[$col->key] = $castedVal;
        }

        if ($options->rowNumberKey !== null) {
            $mapped[$options->rowNumberKey] = $rowNumber;
        }

        return [
            'row' => $mapped,
            'errors' => $errors,
        ];
    }

    /**
     * Check if all cells for the given columns are empty.
     *
     * @param array<int, mixed> $rawRow
     * @param list<ReadColumn> $columns
     */
    public static function isRowEmpty(array $rawRow, array $columns): bool
    {
        foreach ($columns as $col) {
            if ($col->column === null) {
                continue;
            }
            $colIndex = (int) $col->column;
            $val = $rawRow[$colIndex] ?? null;
            if ($val !== null && (!is_string($val) || trim($val) !== '')) {
                return false;
            }
        }

        return true;
    }
}
