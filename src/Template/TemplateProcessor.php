<?php

declare(strict_types=1);

namespace ExcelHelper\Template;

use ExcelHelper\Exceptions\ReaderException;
use ExcelHelper\Exceptions\SheetNotFoundException;
use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\Reader\FileInspector;
use ExcelHelper\Support\Coordinate;
use ExcelHelper\Writer\ExcelOutput;
use ExcelHelper\Writer\WorkbookDefinition;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Throwable;

/**
 * Loads an existing Excel spreadsheet template, replaces cell placeholders,
 * inserts table data while preserving layouts and formatting, and exports the result.
 */
final class TemplateProcessor
{
    private Spreadsheet $spreadsheet;
    private string $format;

    /**
     * @throws ReaderException
     */
    public function __construct(string $templatePath)
    {
        $inspection = FileInspector::inspect($templatePath);
        $this->format = $inspection['format'];

        try {
            $reader = IOFactory::createReader($inspection['readerType']);
            $this->spreadsheet = $reader->load($templatePath);
        } catch (Throwable $e) {
            throw ReaderException::cannotOpenFile($templatePath, $e);
        }
    }

    /**
     * Replace a placeholder like {{key}} or {{ key }} across the worksheet.
     * If the cell content exactly matches {{key}}, the cell value and type will be set directly to $value.
     *
     * @param int|string $sheet Sheet index (0-based) or sheet title
     *
     * @throws SheetNotFoundException
     */
    public function set(string $placeholder, mixed $value, int|string $sheet = 0): self
    {
        $worksheet = $this->resolveSheet($sheet);
        $pattern = '/\{\{\s*' . preg_quote(trim($placeholder, '{} '), '/') . '\s*\}\}/u';

        foreach ($worksheet->getRowIterator() as $row) {
            foreach ($row->getCellIterator() as $cell) {
                $cellVal = $cell->getValue();
                if (!is_string($cellVal) || !preg_match($pattern, $cellVal)) {
                    continue;
                }

                // If cell is exactly {{placeholder}}, replace the whole value with its native type
                if (preg_match('/^\{\{\s*' . preg_quote(trim($placeholder, '{} '), '/') . '\s*\}\}$/uD', trim($cellVal))) {
                    $this->setCellValue($worksheet, $cell->getCoordinate(), $value);
                } else {
                    // Substring replacement
                    $replaced = preg_replace($pattern, (string) $value, $cellVal);
                    $cell->setValueExplicit($replaced ?? '', DataType::TYPE_STRING);
                }
            }
        }

        return $this;
    }

    /**
     * Set multiple placeholders or direct cell addresses (e.g. ['invoice_no' => '123', 'B5' => 'Paid']).
     *
     * @param array<string, mixed> $values
     * @param int|string $sheet Sheet index (0-based) or sheet title
     *
     * @throws SheetNotFoundException
     */
    public function setValues(array $values, int|string $sheet = 0): self
    {
        $worksheet = $this->resolveSheet($sheet);

        foreach ($values as $key => $val) {
            // Direct cell coordinate like 'B5' or 'AA12'
            if (preg_match('/^[A-Z]{1,3}\d+$/D', strtoupper($key))) {
                $this->setCellValue($worksheet, strtoupper($key), $val);
            } else {
                $this->set($key, $val, $sheet);
            }
        }

        return $this;
    }

    /**
     * Insert multiple rows of data at a target placeholder (e.g. '{{items}}') or starting row number.
     * Copies template row formatting, shifts following rows downward, and fills the row data.
     *
     * @param string|int $target Placeholder name (e.g. 'items') or row number (1-based)
     * @param iterable<mixed> $data Rows to insert
     * @param array<array-key, string|array<string, mixed>> $columns Column definitions or mappings
     * @param int|string $sheet Sheet index (0-based) or sheet title
     *
     * @throws SheetNotFoundException
     */
    public function insertTable(
        string|int $target,
        iterable $data,
        array $columns = [],
        int|string $sheet = 0,
    ): self {
        $worksheet = $this->resolveSheet($sheet);
        $dataRows = is_array($data) ? $data : iterator_to_array($data);
        $numRows = count($dataRows);

        // Find starting row and template column mapping
        $startRow = is_int($target) ? $target : $this->findPlaceholderRow($worksheet, $target);
        if ($startRow === null) {
            return $this;
        }

        if ($numRows === 0) {
            // Remove the template placeholder row or clear it
            $worksheet->removeRow($startRow, 1);
            return $this;
        }

        // If more than 1 row, insert additional empty rows with template formatting
        if ($numRows > 1) {
            $worksheet->insertNewRowBefore($startRow + 1, $numRows - 1);
        }

        // Inspect template row cells to copy styles
        $highestCol = $worksheet->getHighestDataColumn($startRow);
        $highestColIdx = Coordinate::columnIndex($highestCol);

        $templateStyles = [];
        $templatePlaceholders = [];
        for ($col = 1; $col <= $highestColIdx; $col++) {
            $addr = Coordinate::formatCell($col, $startRow);
            $cell = $worksheet->getCell($addr);
            $templateStyles[$col] = $worksheet->getStyle($addr);
            $val = $cell->getValue();
            if (is_string($val) && preg_match('/\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}/u', $val, $m)) {
                $templatePlaceholders[$col] = $m[1];
            }
        }

        // Fill data rows
        foreach ($dataRows as $rOffset => $rowItem) {
            $currRow = $startRow + $rOffset;
            /** @var array<array-key, mixed> $rowArr */
            $rowArr = is_array($rowItem) ? $rowItem : (array) $rowItem;
            $isList = array_is_list($rowArr);

            // Copy style to newly inserted row if offset > 0
            if ($rOffset > 0) {
                for ($col = 1; $col <= $highestColIdx; $col++) {
                    $srcAddr = Coordinate::formatCell($col, $startRow);
                    $destAddr = Coordinate::formatCell($col, $currRow);
                    $worksheet->duplicateStyle($worksheet->getStyle($srcAddr), $destAddr);
                }
            }

            // Fill columns
            if ($columns !== []) {
                $cIdx = 1;
                foreach ($columns as $colKey => $colDef) {
                    $keyName = is_string($colDef) ? $colDef : (string) ($colDef['key'] ?? $colKey);
                    $val = $isList ? ($rowArr[$cIdx - 1] ?? null) : ($rowArr[$keyName] ?? null);
                    $cellAddr = Coordinate::formatCell($cIdx, $currRow);
                    $this->setCellValue($worksheet, $cellAddr, $val);
                    $cIdx++;
                }
            } elseif ($templatePlaceholders !== []) {
                // Map based on placeholders found in the template row
                foreach ($templatePlaceholders as $col => $placeholderKey) {
                    $val = $rowArr[$placeholderKey] ?? null;
                    $cellAddr = Coordinate::formatCell($col, $currRow);
                    $this->setCellValue($worksheet, $cellAddr, $val);
                }
            } else {
                // Raw filling across columns
                $cIdx = 1;
                foreach ($rowArr as $val) {
                    $cellAddr = Coordinate::formatCell($cIdx, $currRow);
                    $this->setCellValue($worksheet, $cellAddr, $val);
                    $cIdx++;
                }
            }
        }

        return $this;
    }

    /**
     * Resolve worksheet by index or name.
     *
     * @throws SheetNotFoundException
     */
    public function resolveSheet(int|string $sheet): Worksheet
    {
        $names = array_values($this->spreadsheet->getSheetNames());

        if (is_int($sheet)) {
            $sheetCount = $this->spreadsheet->getSheetCount();
            if ($sheet < 0 || $sheet >= $sheetCount) {
                throw SheetNotFoundException::byIndex($sheet, $names);
            }

            return $this->spreadsheet->getSheet($sheet);
        }

        $ws = $this->spreadsheet->getSheetByName($sheet);
        if ($ws === null) {
            throw SheetNotFoundException::byName($sheet, $names);
        }

        return $ws;
    }

    public function spreadsheet(): Spreadsheet
    {
        return $this->spreadsheet;
    }

    /**
     * Export the processed template into an ExcelOutput wrapper.
     */
    public function output(): ExcelOutput
    {
        return new ExcelOutput(
            spreadsheet: $this->spreadsheet,
            definition: new WorkbookDefinition(format: $this->format),
        );
    }

    /**
     * Save the processed spreadsheet to disk.
     *
     * @throws WriterException
     */
    public function save(string $path): string
    {
        return $this->output()->save($path);
    }

    /**
     * Download the processed spreadsheet.
     *
     * @throws WriterException
     */
    public function download(string $filename): void
    {
        $this->output()->download($filename);
    }

    /**
     * Return binary string of the processed spreadsheet.
     */
    public function toString(): string
    {
        return $this->output()->toString();
    }

    private function findPlaceholderRow(Worksheet $sheet, string $placeholder): ?int
    {
        $clean = trim($placeholder, '{} ');
        $pattern = '/\{\{\s*' . preg_quote($clean, '/') . '\s*\}\}/u';

        foreach ($sheet->getRowIterator() as $row) {
            foreach ($row->getCellIterator() as $cell) {
                $val = $cell->getValue();
                if (is_string($val) && preg_match($pattern, $val)) {
                    return $row->getRowIndex();
                }
            }
        }

        return null;
    }

    private function setCellValue(Worksheet $sheet, string $cellAddress, mixed $value): void
    {
        if ($value === null) {
            $sheet->getCell($cellAddress)->setValueExplicit('', DataType::TYPE_NULL);
            return;
        }

        if (is_bool($value)) {
            $sheet->getCell($cellAddress)->setValueExplicit($value, DataType::TYPE_BOOL);
            return;
        }

        if (is_int($value) || is_float($value)) {
            $sheet->getCell($cellAddress)->setValueExplicit($value, DataType::TYPE_NUMERIC);
            return;
        }

        $str = (string) $value;
        if (str_starts_with($str, '=')) {
            $sheet->getCell($cellAddress)->setValueExplicit($str, DataType::TYPE_FORMULA);
            return;
        }

        $sheet->getCell($cellAddress)->setValueExplicit($str, DataType::TYPE_STRING);
    }
}
