<?php

declare(strict_types=1);

namespace ExcelHelper\Writer;

use ExcelHelper\Config\ConfigPath;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Support\Color;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

/**
 * Maps simple, friendly style configurations into PhpSpreadsheet applyFromArray() structures.
 */
final class StyleMapper
{
    private const ALLOWED_STYLE_KEYS = [
        'bold',
        'italic',
        'underline',
        'strike',
        'color',
        'size',
        'font_size',
        'font',
        'font_name',
        'background',
        'fill',
        'bg',
        'align',
        'alignment',
        'align_horizontal',
        'valign',
        'align_vertical',
        'wrap',
        'wrap_text',
        'borders',
        'format',
        'number_format',
    ];

    private const HORIZONTAL_ALIGNS = [
        'left' => Alignment::HORIZONTAL_LEFT,
        'center' => Alignment::HORIZONTAL_CENTER,
        'right' => Alignment::HORIZONTAL_RIGHT,
        'justify' => Alignment::HORIZONTAL_JUSTIFY,
        'general' => Alignment::HORIZONTAL_GENERAL,
    ];

    private const VERTICAL_ALIGNS = [
        'top' => Alignment::VERTICAL_TOP,
        'center' => Alignment::VERTICAL_CENTER,
        'middle' => Alignment::VERTICAL_CENTER,
        'bottom' => Alignment::VERTICAL_BOTTOM,
        'justify' => Alignment::VERTICAL_JUSTIFY,
    ];

    private const BORDER_STYLES = [
        'none' => Border::BORDER_NONE,
        'thin' => Border::BORDER_THIN,
        'medium' => Border::BORDER_MEDIUM,
        'thick' => Border::BORDER_THICK,
        'dashed' => Border::BORDER_DASHED,
        'dotted' => Border::BORDER_DOTTED,
        'double' => Border::BORDER_DOUBLE,
        'hair' => Border::BORDER_HAIR,
    ];

    /**
     * Map friendly style array to PhpSpreadsheet style array.
     *
     * @param array<string, mixed> $style
     * @return array<string, mixed>
     *
     * @throws InvalidConfigException
     */
    public static function map(array $style, ?ConfigPath $path = null): array
    {
        $path ??= ConfigPath::root('style');
        ConfigPath::assertKnownKeys($style, self::ALLOWED_STYLE_KEYS, $path);

        $result = [];

        // Font options
        $font = [];
        if (isset($style['bold'])) {
            $font['bold'] = (bool) $style['bold'];
        }
        if (isset($style['italic'])) {
            $font['italic'] = (bool) $style['italic'];
        }
        if (isset($style['underline'])) {
            $font['underline'] = (bool) $style['underline'];
        }
        if (isset($style['strike'])) {
            $font['strikethrough'] = (bool) $style['strike'];
        }
        if (isset($style['color'])) {
            $font['color'] = ['rgb' => Color::normalize((string) $style['color'])];
        }

        $fontSize = $style['size'] ?? $style['font_size'] ?? null;
        if ($fontSize !== null) {
            $font['size'] = (float) $fontSize;
        }

        $fontName = $style['font'] ?? $style['font_name'] ?? null;
        if ($fontName !== null) {
            $font['name'] = (string) $fontName;
        }

        if ($font !== []) {
            $result['font'] = $font;
        }

        // Fill / Background options
        $bg = $style['background'] ?? $style['fill'] ?? $style['bg'] ?? null;
        if ($bg !== null) {
            if (is_string($bg)) {
                $result['fill'] = [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['rgb' => Color::normalize($bg)],
                ];
            } elseif (is_array($bg) && isset($bg['color'])) {
                $result['fill'] = [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['rgb' => Color::normalize((string) $bg['color'])],
                ];
            }
        }

        // Alignment options
        $alignment = [];
        $hAlign = $style['align'] ?? $style['alignment'] ?? $style['align_horizontal'] ?? null;
        if (is_string($hAlign)) {
            $normH = strtolower(trim($hAlign));
            if (isset(self::HORIZONTAL_ALIGNS[$normH])) {
                $alignment['horizontal'] = self::HORIZONTAL_ALIGNS[$normH];
            }
        }

        $vAlign = $style['valign'] ?? $style['align_vertical'] ?? null;
        if (is_string($vAlign)) {
            $normV = strtolower(trim($vAlign));
            if (isset(self::VERTICAL_ALIGNS[$normV])) {
                $alignment['vertical'] = self::VERTICAL_ALIGNS[$normV];
            }
        }

        $wrap = $style['wrap'] ?? $style['wrap_text'] ?? null;
        if ($wrap !== null) {
            $alignment['wrapText'] = (bool) $wrap;
        }

        if ($alignment !== []) {
            $result['alignment'] = $alignment;
        }

        // Borders options
        if (isset($style['borders'])) {
            $result['borders'] = self::mapBorders($style['borders']);
        }

        // Number format
        $fmt = $style['format'] ?? $style['number_format'] ?? null;
        if ($fmt !== null) {
            $result['numberFormat'] = ['formatCode' => (string) $fmt];
        }

        return $result;
    }

    /**
     * Map friendly borders to PhpSpreadsheet borders array.
     *
     * @param bool|string|array<string, mixed> $borders
     * @return array<string, mixed>
     */
    public static function mapBorders(bool|string|array $borders): array
    {
        if ($borders === false) {
            return [];
        }

        if ($borders === true || $borders === 'all' || $borders === 'thin') {
            return [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ];
        }

        if ($borders === 'outline') {
            return [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ];
        }

        if (is_string($borders)) {
            $style = self::BORDER_STYLES[strtolower($borders)] ?? Border::BORDER_THIN;

            return [
                'allBorders' => [
                    'borderStyle' => $style,
                    'color' => ['rgb' => '000000'],
                ],
            ];
        }

        $res = [];
        $color = isset($borders['color']) ? Color::normalize((string) $borders['color']) : '000000';
        $style = isset($borders['style']) && isset(self::BORDER_STYLES[strtolower((string) $borders['style'])])
            ? self::BORDER_STYLES[strtolower((string) $borders['style'])]
            : Border::BORDER_THIN;

        $sides = ['top', 'bottom', 'left', 'right', 'allBorders', 'outline', 'inside'];
        $hasSide = false;

        foreach ($sides as $side) {
            if (isset($borders[$side])) {
                $hasSide = true;
                $sideVal = $borders[$side];
                $sideStyle = is_string($sideVal) && isset(self::BORDER_STYLES[strtolower($sideVal)])
                    ? self::BORDER_STYLES[strtolower($sideVal)]
                    : $style;

                $res[$side] = [
                    'borderStyle' => $sideStyle,
                    'color' => ['rgb' => $color],
                ];
            }
        }

        if (!$hasSide) {
            $res['allBorders'] = [
                'borderStyle' => $style,
                'color' => ['rgb' => $color],
            ];
        }

        return $res;
    }
}
