<?php

declare(strict_types=1);

namespace ExcelHelper\Reader;

use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\UnsupportedFormatException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Throwable;

/**
 * Inspects file existence, permissions, size limits, and format whitelist before loading.
 */
final class FileInspector
{
    /**
     * @var array<string, string>
     */
    private const READER_TO_FORMAT = [
        IOFactory::READER_XLSX => 'xlsx',
        IOFactory::READER_XLS => 'xls',
        IOFactory::READER_ODS => 'ods',
        IOFactory::READER_CSV => 'csv',
    ];

    /**
     * @var array<string, string>
     */
    private const FORMAT_TO_READER = [
        'xlsx' => IOFactory::READER_XLSX,
        'xls' => IOFactory::READER_XLS,
        'ods' => IOFactory::READER_ODS,
        'csv' => IOFactory::READER_CSV,
    ];

    /**
     * Inspect file and determine the appropriate PhpSpreadsheet reader type and normalized format.
     *
     * @return array{readerType: string, format: string}
     *
     * @throws ReaderException
     * @throws UnsupportedFormatException
     */
    public static function inspect(string $path, ReadOptions $options): array
    {
        if (!file_exists($path)) {
            throw ReaderException::fileNotFound($path);
        }

        if (is_dir($path)) {
            throw ReaderException::isADirectory($path);
        }

        if (!is_readable($path)) {
            throw ReaderException::notReadable($path);
        }

        $size = filesize($path);
        if ($size === false || $size === 0) {
            throw ReaderException::emptyFile($path);
        }

        if ($options->maxFileSize !== null && $size > $options->maxFileSize) {
            throw ReaderException::maxFileSizeExceeded($path, $options->maxFileSize, $size);
        }

        $declaredFormat = strtolower(trim($options->format));

        if ($declaredFormat !== 'auto') {
            if (!in_array($declaredFormat, $options->allowedFormats, true)) {
                throw UnsupportedFormatException::notWhitelisted($declaredFormat, $options->allowedFormats);
            }

            if (!isset(self::FORMAT_TO_READER[$declaredFormat])) {
                throw UnsupportedFormatException::unsupportedForReading($declaredFormat);
            }

            $readerType = self::FORMAT_TO_READER[$declaredFormat];
            $format = $declaredFormat;
        } else {
            try {
                $detectedReader = IOFactory::identify($path);
            } catch (Throwable) {
                throw UnsupportedFormatException::unrecognized($path);
            }

            $format = self::READER_TO_FORMAT[$detectedReader] ?? strtolower($detectedReader);
            $readerType = $detectedReader;

            if (!in_array($format, $options->allowedFormats, true)) {
                throw UnsupportedFormatException::notWhitelisted($format, $options->allowedFormats);
            }
        }

        return [
            'readerType' => $readerType,
            'format' => $format,
        ];
    }
}
