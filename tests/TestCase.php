<?php

declare(strict_types=1);

namespace ExcelHelper\Tests;

use ExcelHelper\Excel;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate as PhpSpreadsheetCoordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\TestCase as BaseTestCase;
use ZipArchive;

abstract class TestCase extends BaseTestCase
{
    /**
     * @var list<string>
     */
    private array $tempFiles = [];

    protected function tearDown(): void
    {
        Excel::resetDefaults();

        foreach ($this->tempFiles as $file) {
            if (is_file($file)) {
                @unlink($file);
            } elseif (is_dir($file)) {
                $this->removeDirectory($file);
            }
        }

        $this->tempFiles = [];

        parent::tearDown();
    }

    /**
     * Generate a unique temporary file path and register it for cleanup.
     */
    protected function tempPath(string $extension = 'xlsx'): string
    {
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'excel_helper_tests';
        if (!is_dir($dir)) {
            mkdir($dir, 0777, true);
        }

        $ext = ltrim($extension, '.');
        $path = $dir . DIRECTORY_SEPARATOR . 'test_' . bin2hex(random_bytes(8)) . ($ext !== '' ? '.' . $ext : '');
        $this->tempFiles[] = $path;

        return $path;
    }

    /**
     * Create a dynamic spreadsheet file with raw rows and return its path.
     *
     * @param list<list<mixed>>|list<array<string, mixed>> $rows
     * @param array<string, mixed> $options
     */
    protected function createSpreadsheetFile(
        array $rows,
        string $format = 'xlsx',
        string $sheetName = 'Sheet1',
        array $options = [],
    ): string {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle($sheetName);

        $rowIdx = 1;
        foreach ($rows as $row) {
            $colIdx = 1;
            foreach ($row as $val) {
                $cellAddress = PhpSpreadsheetCoordinate::stringFromColumnIndex($colIdx) . $rowIdx;
                $sheet->setCellValue($cellAddress, $val);
                $colIdx++;
            }
            $rowIdx++;
        }

        if (isset($options['date_system_1904']) && $options['date_system_1904'] === true) {
            \PhpOffice\PhpSpreadsheet\Shared\Date::setExcelCalendar(\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_MAC_1904);
        }

        $path = $this->tempPath($format);
        $writerType = match (strtolower($format)) {
            'xlsx' => 'Xlsx',
            'xls' => 'Xls',
            'ods' => 'Ods',
            'csv' => 'Csv',
            default => 'Xlsx',
        };

        $writer = IOFactory::createWriter($spreadsheet, $writerType);
        $writer->save($path);
        $spreadsheet->disconnectWorksheets();

        // Reset calendar back to default 1900
        \PhpOffice\PhpSpreadsheet\Shared\Date::setExcelCalendar(\PhpOffice\PhpSpreadsheet\Shared\Date::CALENDAR_WINDOWS_1900);

        return $path;
    }

    /**
     * Create a multi-sheet spreadsheet file.
     *
     * @param list<array{name: string, rows: list<list<mixed>>}> $sheets
     */
    protected function createMultiSheetFile(array $sheets, string $format = 'xlsx'): string
    {
        $spreadsheet = new Spreadsheet();

        foreach ($sheets as $index => $sheetData) {
            $worksheet = $index === 0 ? $spreadsheet->getActiveSheet() : $spreadsheet->createSheet();
            $worksheet->setTitle($sheetData['name']);

            $rowIdx = 1;
            foreach ($sheetData['rows'] as $row) {
                $colIdx = 1;
                foreach ($row as $val) {
                    $cellAddress = PhpSpreadsheetCoordinate::stringFromColumnIndex($colIdx) . $rowIdx;
                    $worksheet->setCellValue($cellAddress, $val);
                    $colIdx++;
                }
                $rowIdx++;
            }
        }

        $path = $this->tempPath($format);
        $writerType = match (strtolower($format)) {
            'xlsx' => 'Xlsx',
            'xls' => 'Xls',
            'ods' => 'Ods',
            'csv' => 'Csv',
            default => 'Xlsx',
        };

        $writer = IOFactory::createWriter($spreadsheet, $writerType);
        $writer->save($path);
        $spreadsheet->disconnectWorksheets();

        return $path;
    }

    /**
     * Create a CSV file with specific content and encoding.
     */
    protected function createCsvFile(string $content, string $encoding = 'UTF-8'): string
    {
        $path = $this->tempPath('csv');

        $encodedContent = $encoding === 'UTF-8'
            ? $content
            : mb_convert_encoding($content, $encoding, 'UTF-8');

        file_put_contents($path, $encodedContent);

        return $path;
    }

    /**
     * Create a fake spreadsheet file (e.g. HTML or text with .xlsx extension).
     */
    protected function createFakeSpreadsheet(string $extension = 'xlsx'): string
    {
        $path = $this->tempPath($extension);
        file_put_contents($path, '<html><body><h1>Not a spreadsheet</h1><script>alert("xss")</script></body></html>');

        return $path;
    }

    /**
     * Create a malicious XLSX file containing an XXE payload in XML structure.
     */
    protected function createXxeXlsx(): string
    {
        $validFile = $this->createSpreadsheetFile([['Header'], ['Value']], 'xlsx');
        $maliciousPath = $this->tempPath('xlsx');

        copy($validFile, $maliciousPath);

        $zip = new ZipArchive();
        if ($zip->open($maliciousPath) === true) {
            $xxeXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                . '<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>'
                . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                . '<sheets><sheet name="Sheet1&xxe;" sheetId="1" r:id="rId1"/></sheets>'
                . '</workbook>';
            $zip->addFromString('xl/workbook.xml', $xxeXml);
            $zip->close();
        }

        return $maliciousPath;
    }

    private function removeDirectory(string $dir): void
    {
        $items = scandir($dir);
        if ($items === false) {
            return;
        }

        foreach ($items as $item) {
            if ($item === '.' || $item === '..') {
                continue;
            }
            $subPath = $dir . DIRECTORY_SEPARATOR . $item;
            if (is_dir($subPath)) {
                $this->removeDirectory($subPath);
            } else {
                @unlink($subPath);
            }
        }

        @rmdir($dir);
    }
}
