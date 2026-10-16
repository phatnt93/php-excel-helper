<?php

declare(strict_types=1);

namespace ExcelHelper\Support;

/**
 * Generates secure HTTP download headers preventing header injection and handling UTF-8 filenames via RFC 5987.
 */
final class DownloadHeaders
{
    private const MIME_TYPES = [
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'xls' => 'application/vnd.ms-excel',
        'ods' => 'application/vnd.oasis.opendocument.spreadsheet',
        'csv' => 'text/csv; charset=UTF-8',
    ];

    /**
     * @return array<string, string>
     */
    public static function create(string $filename, string $format): array
    {
        // Sanitize filename against CRLF header injection
        $cleanName = str_replace(["\r", "\n", "\0", '"', '\\'], '', trim($filename));
        if ($cleanName === '') {
            $cleanName = 'download.' . $format;
        }

        $replaced = preg_replace('/[^A-Za-z0-9._-]+/', '_', $cleanName);
        $asciiFallback = is_string($replaced) ? $replaced : 'download.' . $format;

        $encodedName = rawurlencode($cleanName);

        $contentType = self::MIME_TYPES[strtolower($format)] ?? 'application/octet-stream';
        $disposition = sprintf('attachment; filename="%s"; filename*=UTF-8\'\'%s', $asciiFallback, $encodedName);

        return [
            'Content-Type' => $contentType,
            'Content-Disposition' => $disposition,
            'Cache-Control' => 'max-age=0, no-cache, no-store, must-revalidate',
            'Pragma' => 'public',
            'Expires' => '0',
        ];
    }
}
