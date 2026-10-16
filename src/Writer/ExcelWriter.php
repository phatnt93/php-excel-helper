<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Builds a PhpSpreadsheet instance from a WorkbookDefinition.
 */
final class ExcelWriter
{
    public static function build(
        WorkbookDefinition $definition,
        ?Spreadsheet $existingSpreadsheet = null,
    ): Spreadsheet {
        $spreadsheet = $existingSpreadsheet ?? new Spreadsheet();

        if ($definition->properties !== []) {
            $props = $spreadsheet->getProperties();
            if (isset($definition->properties['creator'])) {
                $props->setCreator((string) $definition->properties['creator']);
            }
            if (isset($definition->properties['title'])) {
                $props->setTitle((string) $definition->properties['title']);
            }
            if (isset($definition->properties['subject'])) {
                $props->setSubject((string) $definition->properties['subject']);
            }
            if (isset($definition->properties['description'])) {
                $props->setDescription((string) $definition->properties['description']);
            }
            if (isset($definition->properties['company'])) {
                $props->setCompany((string) $definition->properties['company']);
            }
            if (isset($definition->properties['category'])) {
                $props->setCategory((string) $definition->properties['category']);
            }
            if (isset($definition->properties['keywords'])) {
                $props->setKeywords((string) $definition->properties['keywords']);
            }
        }

        $existingSheetNames = [];
        foreach ($definition->sheets as $sheetDef) {
            $layout = SheetRenderer::render(
                spreadsheet: $spreadsheet,
                definition: $sheetDef,
                format: $definition->format,
                existingSheetNames: $existingSheetNames,
                workbookCsvOptions: $definition->csv,
                locale: $definition->locale,
            );
            $existingSheetNames[] = $sheetDef->name;
        }

        return $spreadsheet;
    }
}
