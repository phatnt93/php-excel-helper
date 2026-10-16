<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use DateTime;
use DateTimeImmutable;
use DateTimeInterface;
use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;
use PhpOffice\PhpSpreadsheet\Shared\Date as PhpSpreadsheetDate;
use Throwable;

/**
 * DateTime column type.
 */
final class DateTimeType implements TypeInterface
{
    /**
     * @var list<string>
     */
    private const COMMON_FORMATS = [
        'Y-m-d H:i:s',
        'Y-m-d H:i',
        'd/m/Y H:i:s',
        'd/m/Y H:i',
        'd-m-Y H:i:s',
        'Y-m-d',
        'd/m/Y',
    ];

    public function name(): string
    {
        return 'datetime';
    }

    public function read(mixed $value, ColumnContext $context): string|DateTimeImmutable
    {
        $dt = $this->parseToDateTime($value, $context);

        $outFormat = $context->getDateTimeFormat();
        if ($outFormat === null) {
            return $dt;
        }

        return $dt->format($outFormat);
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        $dt = $this->toDateTimeInterface($value, $context);
        $excelDate = PhpSpreadsheetDate::PHPToExcel($dt);
        if ($excelDate === false) {
            throw WriterException::cannotWriteValue($context->key, 'Failed to convert datetime to Excel serial value.');
        }
        /** @var string $formatCode */
        $formatCode = $context->get('excel_format', 'yyyy-mm-dd hh:mm:ss');

        return CellValue::numeric($excelDate, $formatCode);
    }

    private function parseToDateTime(mixed $value, ColumnContext $context): DateTimeImmutable
    {
        if ($value instanceof DateTimeImmutable) {
            return $value;
        }

        if ($value instanceof DateTimeInterface) {
            return DateTimeImmutable::createFromInterface($value);
        }

        if (is_int($value) || is_float($value)) {
            $calendar = $context->isDateSystem1904()
                ? PhpSpreadsheetDate::CALENDAR_MAC_1904
                : PhpSpreadsheetDate::CALENDAR_WINDOWS_1900;

            try {
                $dt = PhpSpreadsheetDate::excelToDateTimeObject((float) $value, null, $calendar);

                return DateTimeImmutable::createFromMutable($dt);
            } catch (Throwable) {
                throw new TypeCastException('invalid_date', ['value' => $value]);
            }
        }

        if (is_string($value)) {
            $str = $context->isTrim() ? trim($value) : $value;
            if ($str === '') {
                throw new TypeCastException('invalid_date', ['value' => $value]);
            }

            /** @var string|list<string>|null $inputFormat */
            $inputFormat = $context->get('input_format');
            $formats = $inputFormat !== null ? (is_array($inputFormat) ? $inputFormat : [$inputFormat]) : self::COMMON_FORMATS;

            foreach ($formats as $fmt) {
                $dt = DateTimeImmutable::createFromFormat('!' . $fmt, $str);
                if ($dt !== false) {
                    $errors = DateTimeImmutable::getLastErrors();
                    if ($errors === false || ($errors['warning_count'] === 0 && $errors['error_count'] === 0)) {
                        return $dt;
                    }
                }
            }
        }

        throw new TypeCastException('invalid_date', ['value' => $value]);
    }

    private function toDateTimeInterface(mixed $value, ColumnContext $context): DateTimeInterface
    {
        if ($value instanceof DateTimeInterface) {
            return $value;
        }

        if (is_int($value)) {
            return (new DateTime())->setTimestamp($value);
        }

        if (is_string($value)) {
            try {
                return $this->parseToDateTime($value, $context);
            } catch (TypeCastException) {
                throw WriterException::cannotWriteValue($context->key, sprintf('Value "%s" cannot be parsed as a datetime.', $value));
            }
        }

        throw WriterException::cannotWriteValue($context->key, 'Unsupported datetime value type.');
    }
}
