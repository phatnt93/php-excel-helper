<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Support;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Support\Color;
use ExcelHelper\Tests\TestCase;

final class ColorTest extends TestCase
{
    public function testNormalize3Hex(): void
    {
        $this->assertSame('FFFFFF', Color::normalize('#FFF'));
        $this->assertSame('FFFFFF', Color::normalize('FFF'));
        $this->assertSame('112233', Color::normalize('#123'));
        $this->assertSame('AABBCC', Color::normalize('abc'));
    }

    public function testNormalize6Hex(): void
    {
        $this->assertSame('1F4E78', Color::normalize('#1F4E78'));
        $this->assertSame('1F4E78', Color::normalize('1f4e78'));
        $this->assertSame('FFFFFF', Color::normalize('  #ffffff  '));
    }

    public function testNormalizeInvalidThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Invalid hex color format');
        Color::normalize('#GGGGGG');
    }

    public function testNormalizeTooShortThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        Color::normalize('#12');
    }

    public function testNormalizeTooLongThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        Color::normalize('#1234567');
    }

    public function testNormalizeEmptyThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        Color::normalize('');
    }
}
