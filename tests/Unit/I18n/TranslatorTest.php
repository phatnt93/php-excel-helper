<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\I18n;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\I18n\Translator;
use ExcelHelper\Tests\TestCase;

final class TranslatorTest extends TestCase
{
    public function testDefaultLocale(): void
    {
        $this->assertSame('en', Translator::getLocale());
    }

    public function testSetLocaleAndReset(): void
    {
        Translator::setLocale('vi');
        $this->assertSame('vi', Translator::getLocale());

        Translator::reset();
        $this->assertSame('en', Translator::getLocale());
    }

    public function testSetUnsupportedLocaleThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Unsupported locale "fr"');
        Translator::setLocale('fr');
    }

    public function testTranslateEnglish(): void
    {
        $message = Translator::trans('row.required', [
            'row' => 5,
            'column' => 'B',
            'key' => 'email',
        ], 'en');

        $this->assertSame('Row 5, column B (email): Value is required.', $message);
    }

    public function testTranslateVietnamese(): void
    {
        $message = Translator::trans('row.required', [
            'row' => 5,
            'column' => 'B',
            'key' => 'email',
        ], 'vi');

        $this->assertSame('Dòng 5, cột B (email): Giá trị là bắt buộc.', $message);
    }

    public function testTranslateInterpolatesTypes(): void
    {
        $message = Translator::trans('row.invalid_bool', [
            'row' => 2,
            'column' => 'A',
            'key' => 'active',
            'value' => 'maybe',
        ], 'en');

        $this->assertSame('Row 2, column A (active): Value "maybe" is not a valid boolean.', $message);
    }

    public function testTranslateMissingKeyReturnsKey(): void
    {
        $key = 'nonexistent.key';
        $this->assertSame($key, Translator::trans($key));
    }
}
