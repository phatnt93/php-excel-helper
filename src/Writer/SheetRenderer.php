<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

use ExcelHelper\Exceptions\WriterException;
use ExcelHelper\I18n\Translator;
use ExcelHelper\Support\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Renders a SheetDefinition onto a PhpSpreadsheet Worksheet with titles, headers, data, totals, and styles.
 */
final class SheetRenderer
{
    private const XLS_MAX_ROWS = 65536;
    private const XLS_MAX_COLUMNS = 256;

    /**
     * Render the given SheetDefinition onto a worksheet in the spreadsheet.
     *
     * @param list<string> $existingSheetNames
     * @param array<string, mixed> $workbookCsvOptions
     *
     * @throws WriterException
     */
    public static function render(
        Spreadsheet $spreadsheet,
        SheetDefinition $definition,
        string $format,
        array $existingSheetNames = [],
        array $workbookCsvOptions = [],
        ?string $locale = null,
    ): SheetLayout {
        $worksheet = $spreadsheet->getActiveSheetIndex() === 0 && $spreadsheet->getSheetCount() === 1 && $existingSheetNames === []
            ? $spreadsheet->getActiveSheet()
            : $spreadsheet->createSheet();

        $sheetName = SheetNameSanitizer::sanitize($definition->name, $existingSheetNames);
        $worksheet->setTitle($sheetName);

        // Parse starting coordinate
        $startPos = Coordinate::parseCell($definition->startCell);
        $startCol = $startPos['col'];
        $currRow = $startPos['row'];

        // Resolve columns
        $columns = self::resolveColumns($definition);
        $numColumns = count($columns);
        $endCol = $numColumns > 0 ? $startCol + $numColumns - 1 : $startCol;

        if ($format === 'xls' && $endCol > self::XLS_MAX_COLUMNS) {
            throw WriterException::xlsLimitExceeded(1, $endCol, self::XLS_MAX_ROWS, self::XLS_MAX_COLUMNS);
        }

        // 1. Render Title banner
        $titleCell = null;
        if ($definition->title !== null && $definition->title !== '') {
            $titleCell = Coordinate::formatCell($startCol, $currRow);
            $titleEndCell = Coordinate::formatCell($endCol, $currRow);

            $worksheet->setCellValueExplicit($titleCell, $definition->title, DataType::TYPE_STRING);

            if ($endCol > $startCol) {
                $worksheet->mergeCells($titleCell . ':' . $titleEndCell);
            }

            $titleStyle = $definition->titleStyle !== null
                ? StyleMapper::map($definition->titleStyle)
                : [
                    'font' => ['bold' => true, 'size' => 14],
                    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
                ];

            $worksheet->getStyle($titleCell . ':' . $titleEndCell)->applyFromArray($titleStyle);
            $currRow += 2; // Move down past title and spacing
        }

        // 2. Render Header row
        $headerRow = null;
        if ($definition->showHeader) {
            $headerRow = $currRow;
            $headerRange = Coordinate::formatCell($startCol, $headerRow) . ':' . Coordinate::formatCell($endCol, $headerRow);

            foreach ($columns as $idx => $col) {
                $cIdx = $startCol + $idx;
                $cellAddr = Coordinate::formatCell($cIdx, $headerRow);
                $headerText = $col->header ?? $col->key;
                $worksheet->setCellValueExplicit($cellAddr, $headerText, DataType::TYPE_STRING);
            }

            $headerStyle = $definition->headerStyle !== null
                ? StyleMapper::map($definition->headerStyle)
                : [
                    'font' => ['bold' => true, 'color' => ['rgb' => '000000']],
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => 'E0E0E0'],
                    ],
                    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
                ];

            $worksheet->getStyle($headerRange)->applyFromArray($headerStyle);

            if ($definition->freezeHeader) {
                $freezeCell = Coordinate::formatCell($startCol, $headerRow + 1);
                $worksheet->freezePane($freezeCell);
            }

            if ($definition->autoFilter && $numColumns > 0) {
                $worksheet->setAutoFilter($headerRange);
            }

            $currRow++;
        }

        // 3. Render Data rows
        $dataStartRow = $currRow;
        $columnNumberFormats = [];
        $rowCount = 0;

        foreach ($definition->data as $rowData) {
            if ($format === 'xls' && $currRow > self::XLS_MAX_ROWS) {
                throw WriterException::xlsLimitExceeded($currRow, $endCol, self::XLS_MAX_ROWS, self::XLS_MAX_COLUMNS);
            }

            /** @var array<array-key, mixed> $assocRow */
            $assocRow = is_array($rowData) ? $rowData : (array) $rowData;
            $isListRow = array_is_list($assocRow);

            foreach ($columns as $idx => $col) {
                $cIdx = $startCol + $idx;
                $cellAddr = Coordinate::formatCell($cIdx, $currRow);

                $val = null;
                if ($isListRow) {
                    $val = $assocRow[$idx] ?? null;
                } else {
                    $val = $assocRow[$col->key] ?? null;
                }

                if (
                    is_string($val)
                    && $format === 'csv'
                    && ($workbookCsvOptions['escape_formulas'] ?? true)
                    && $col->type !== 'formula'
                    && !is_numeric($val)
                ) {
                    if ($val !== '' && str_contains("=+-@\t\r", $val[0])) {
                        $val = "'" . $val;
                    }
                }

                $formatCode = CellWriter::writeCell($worksheet, $cellAddr, $val, $col, $assocRow);
                if ($formatCode !== null && !isset($columnNumberFormats[$idx])) {
                    $columnNumberFormats[$idx] = $formatCode;
                }
            }

            // Zebra styling
            if ($definition->zebra && $rowCount % 2 === 1) {
                $rowRange = Coordinate::formatCell($startCol, $currRow) . ':' . Coordinate::formatCell($endCol, $currRow);
                $zebraColor = is_array($definition->zebra) && isset($definition->zebra['color'])
                    ? \ExcelHelper\Support\Color::normalize((string) $definition->zebra['color'])
                    : 'F9F9F9';

                $worksheet->getStyle($rowRange)->applyFromArray([
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => $zebraColor],
                    ],
                ]);
            }

            if ($definition->rowStyle !== null) {
                $rowRange = Coordinate::formatCell($startCol, $currRow) . ':' . Coordinate::formatCell($endCol, $currRow);
                $worksheet->getStyle($rowRange)->applyFromArray(StyleMapper::map($definition->rowStyle));
            }

            $currRow++;
            $rowCount++;
        }

        $dataEndRow = $rowCount > 0 ? $currRow - 1 : $dataStartRow;

        // Apply number formats by column ranges in bulk for performance
        if ($rowCount > 0) {
            foreach ($columnNumberFormats as $idx => $fmtCode) {
                $cIdx = $startCol + $idx;
                $colRange = Coordinate::formatCell($cIdx, $dataStartRow) . ':' . Coordinate::formatCell($cIdx, $dataEndRow);
                $worksheet->getStyle($colRange)->getNumberFormat()->setFormatCode($fmtCode);
            }
        }

        // Apply column-specific styles
        if ($rowCount > 0) {
            foreach ($columns as $idx => $col) {
                if ($col->style !== null) {
                    $cIdx = $startCol + $idx;
                    $colRange = Coordinate::formatCell($cIdx, $dataStartRow) . ':' . Coordinate::formatCell($cIdx, $dataEndRow);
                    $worksheet->getStyle($colRange)->applyFromArray(StyleMapper::map($col->style));
                }
            }
        }

        // 4. Render Totals row
        $totalsRow = null;
        if ($definition->totals !== null && $definition->totals !== [] && $numColumns > 0) {
            $totalsRow = $currRow;
            $totalsRange = Coordinate::formatCell($startCol, $totalsRow) . ':' . Coordinate::formatCell($endCol, $totalsRow);

            $totalsLabel = $definition->totalsLabel ?? Translator::trans('label.totals', [], $locale);
            $worksheet->setCellValueExplicit(
                Coordinate::formatCell($startCol, $totalsRow),
                $totalsLabel,
                DataType::TYPE_STRING,
            );

            foreach ($columns as $idx => $col) {
                if (isset($definition->totals[$col->key])) {
                    $func = strtoupper($definition->totals[$col->key]);
                    $cIdx = $startCol + $idx;
                    $colLetter = Coordinate::columnLetter($cIdx);
                    $cellAddr = Coordinate::formatCell($cIdx, $totalsRow);

                    $formula = $rowCount > 0
                        ? sprintf('=%s(%s%d:%s%d)', $func, $colLetter, $dataStartRow, $colLetter, $dataEndRow)
                        : sprintf('=%s(0)', $func);

                    $worksheet->setCellValueExplicit($cellAddr, $formula, DataType::TYPE_FORMULA);

                    // Re-apply number format from column if any
                    if (isset($columnNumberFormats[$idx])) {
                        $worksheet->getStyle($cellAddr)->getNumberFormat()->setFormatCode($columnNumberFormats[$idx]);
                    }
                }
            }

            $worksheet->getStyle($totalsRange)->applyFromArray([
                'font' => ['bold' => true],
                'borders' => [
                    'top' => ['borderStyle' => Border::BORDER_THIN],
                    'bottom' => ['borderStyle' => Border::BORDER_DOUBLE],
                ],
            ]);

            $currRow++;
        }

        // 5. Borders styling for entire table
        $tableStartRow = $headerRow ?? $dataStartRow;
        $tableEndRow = $totalsRow ?? $dataEndRow;
        $tableRange = Coordinate::formatCell($startCol, $tableStartRow) . ':' . Coordinate::formatCell($endCol, $tableEndRow);

        if ($definition->borders !== false) {
            $borderStyle = StyleMapper::mapBorders($definition->borders);

            if ($borderStyle !== []) {
                $worksheet->getStyle($tableRange)->applyFromArray(['borders' => $borderStyle]);
            }
        }

        // 6. Layout: Auto-size and Column widths
        $shouldAutoSize = $definition->autoSize ?? ($rowCount <= $definition->autoSizeMaxRows);

        foreach ($columns as $idx => $col) {
            $cIdx = $startCol + $idx;
            $colLetter = Coordinate::columnLetter($cIdx);
            $dimension = $worksheet->getColumnDimension($colLetter);

            if ($col->width !== null) {
                if ($col->width === 'auto') {
                    $dimension->setAutoSize(true);
                } else {
                    $dimension->setWidth((float) $col->width);
                }
            } elseif ($shouldAutoSize) {
                $dimension->setAutoSize(true);
            }
        }

        if ($definition->rowHeight !== null) {
            for ($r = $tableStartRow; $r <= $tableEndRow; $r++) {
                $worksheet->getRowDimension($r)->setRowHeight($definition->rowHeight);
            }
        }

        // Custom merges
        foreach ($definition->merges as $mergeRange) {
            $worksheet->mergeCells($mergeRange);
        }

        $dataRange = Coordinate::formatCell($startCol, $dataStartRow) . ':' . Coordinate::formatCell($endCol, $dataEndRow);

        $layout = new SheetLayout(
            titleCell: $titleCell,
            headerRow: $headerRow,
            dataStartRow: $dataStartRow,
            dataEndRow: $dataEndRow,
            totalsRow: $totalsRow,
            startCol: $startCol,
            endCol: $endCol,
            dataRange: $dataRange,
            tableRange: $tableRange,
        );

        // 7. Tap hook
        if ($definition->tap !== null) {
            ($definition->tap)($worksheet, $layout);
        }

        return $layout;
    }

    /**
     * @return list<WriteColumn>
     */
    private static function resolveColumns(SheetDefinition $def): array
    {
        if ($def->columns !== []) {
            return $def->columns;
        }

        if ($def->headers !== null && $def->headers !== []) {
            $cols = [];
            foreach ($def->headers as $idx => $headerText) {
                $cols[] = new WriteColumn(
                    key: (string) $idx,
                    header: $headerText,
                    type: 'auto',
                );
            }

            return $cols;
        }

        // Try to deduce from first row of data
        foreach ($def->data as $firstRow) {
            if (is_array($firstRow)) {
                $cols = [];
                $isList = array_is_list($firstRow);
                foreach (array_keys($firstRow) as $k) {
                    $cols[] = new WriteColumn(
                        key: (string) $k,
                        header: $isList ? Coordinate::columnLetter((int) $k + 1) : (string) $k,
                        type: 'auto',
                    );
                }

                return $cols;
            }
            break;
        }

        return [];
    }
}
