<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\Support\DownloadHeaders;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use Throwable;

/**
 * Output wrapper providing save(), toString(), download(), and spreadsheet() access.
 */
final class ExcelOutput
{
    private const FORMAT_MAP = [
        'xlsx' => IOFactory::WRITER_XLSX,
        'xls' => IOFactory::WRITER_XLS,
        'ods' => IOFactory::WRITER_ODS,
        'csv' => IOFactory::WRITER_CSV,
    ];

    public function __construct(
        private readonly Spreadsheet $spreadsheet,
        private readonly WorkbookDefinition $definition,
    ) {}

    /**
     * Save the spreadsheet to a local file atomically.
     *
     * @throws WriterException
     */
    public function save(string $path): string
    {
        $dir = dirname($path);
        if (!is_dir($dir) || !is_writable($dir)) {
            throw WriterException::directoryNotWritable($dir);
        }

        $format = $this->resolveFormatFromPath($path);
        $writer = $this->createWriter($format);

        $tempPath = $dir . DIRECTORY_SEPARATOR . uniqid('.tmp_excel_', true) . '.' . $format;

        try {
            $writer->save($tempPath);
            if (!@rename($tempPath, $path)) {
                @unlink($tempPath);
                throw WriterException::failedToSave($path);
            }
        } catch (Throwable $e) {
            if (file_exists($tempPath)) {
                @unlink($tempPath);
            }
            if ($e instanceof WriterException) {
                throw $e;
            }

            throw WriterException::failedToSave($path, $e);
        }

        return realpath($path) ?: $path;
    }

    /**
     * Return binary content of the exported file.
     *
     * @throws WriterException
     */
    public function toString(): string
    {
        $format = $this->definition->format;
        $writer = $this->createWriter($format);

        $tempFile = tempnam(sys_get_temp_dir(), 'excel_out_');
        if ($tempFile === false) {
            throw WriterException::withMessage('Failed to create temporary file for spreadsheet export.');
        }

        try {
            $writer->save($tempFile);
            $content = file_get_contents($tempFile);
            if ($content === false) {
                throw WriterException::withMessage('Failed to read exported spreadsheet content.');
            }

            return $content;
        } catch (Throwable $e) {
            if ($e instanceof WriterException) {
                throw $e;
            }

            throw WriterException::withMessage('Error generating spreadsheet binary string.', $e);
        } finally {
            if (file_exists($tempFile)) {
                @unlink($tempFile);
            }
        }
    }

    /**
     * Send HTTP headers and stream download to browser.
     *
     * @throws WriterException
     */
    public function download(string $filename): void
    {
        if (headers_sent()) {
            throw WriterException::headersAlreadySent();
        }

        $format = $this->resolveFormatFromPath($filename);
        $headers = DownloadHeaders::create($filename, $format);

        foreach ($headers as $headerName => $headerValue) {
            header($headerName . ': ' . $headerValue);
        }

        $writer = $this->createWriter($format);
        $writer->save('php://output');
    }

    /**
     * Access the underlying PhpSpreadsheet instance directly.
     */
    public function spreadsheet(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    private function resolveFormatFromPath(string $path): string
    {
        $ext = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        if (in_array($ext, WorkbookDefinition::SUPPORTED_FORMATS, true)) {
            return $ext;
        }

        return $this->definition->format;
    }

    private function createWriter(string $format): IWriter
    {
        $writerType = self::FORMAT_MAP[$format] ?? IOFactory::WRITER_XLSX;
        $writer = IOFactory::createWriter($this->spreadsheet, $writerType);

        if ($writer instanceof CsvWriter) {
            $csv = $this->definition->csv;
            if (isset($csv['delimiter']) && is_string($csv['delimiter'])) {
                $writer->setDelimiter($csv['delimiter']);
            }
            if (isset($csv['enclosure']) && is_string($csv['enclosure'])) {
                $writer->setEnclosure($csv['enclosure']);
            }
            if (isset($csv['line_ending']) && is_string($csv['line_ending'])) {
                $writer->setLineEnding($csv['line_ending']);
            }

            // BOM default is true for CSV
            $useBom = (bool) ($csv['bom'] ?? true);
            $writer->setUseBOM($useBom);
        }

        return $writer;
    }
}
