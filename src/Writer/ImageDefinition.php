<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

/**
 * Immutable value object holding configuration for an image or logo inserted into a worksheet.
 */
final class ImageDefinition
{
    public function __construct(
        public readonly string $path,
        public readonly string $cell = 'A1',
        public readonly ?int $width = null,
        public readonly ?int $height = null,
        public readonly int $offsetX = 0,
        public readonly int $offsetY = 0,
        public readonly ?string $name = null,
        public readonly ?string $description = null,
    ) {}
}
