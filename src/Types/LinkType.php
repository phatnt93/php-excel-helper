<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\WriterException;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

/**
 * Hyperlink column type.
 */
final class LinkType implements TypeInterface
{
    private const ALLOWED_SCHEMES = ['http', 'https', 'mailto'];

    public function name(): string
    {
        return 'link';
    }

    public function read(mixed $value, ColumnContext $context): string
    {
        $str = (string) $value;

        return $context->isTrim() ? trim($str) : $str;
    }

    public function write(mixed $value, ColumnContext $context): CellValue
    {
        if ($value === null) {
            return CellValue::null();
        }

        $url = '';
        $label = '';

        if (is_array($value)) {
            $url = isset($value['url']) ? (string) $value['url'] : '';
            $label = isset($value['label']) ? (string) $value['label'] : $url;
        } elseif (is_string($value)) {
            $url = $value;
            $label = $value;
        } else {
            throw WriterException::cannotWriteValue($context->key, 'Link value must be a URL string or an array with "url" and "label".');
        }

        $trimmedUrl = trim($url);
        $scheme = parse_url($trimmedUrl, PHP_URL_SCHEME);

        if (is_string($scheme)) {
            $lowerScheme = strtolower($scheme);
            if (!in_array($lowerScheme, self::ALLOWED_SCHEMES, true)) {
                throw WriterException::cannotWriteValue(
                    $context->key,
                    sprintf('Forbidden link scheme "%s:". Only http, https, and mailto schemes are allowed.', $lowerScheme),
                );
            }
        } elseif (!str_starts_with($trimmedUrl, '/') && !str_starts_with($trimmedUrl, './')) {
            throw WriterException::cannotWriteValue(
                $context->key,
                sprintf('Invalid link URL "%s". Expected an absolute URL or relative path.', $url),
            );
        }

        return new CellValue($label, DataType::TYPE_STRING, null, $trimmedUrl);
    }
}
