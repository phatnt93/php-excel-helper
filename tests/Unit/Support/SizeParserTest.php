<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Support;

use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Support\SizeParser;
use ExcelHelper\Tests\TestCase;

final class SizeParserTest extends TestCase
{
    public function testParseNullAndEmpty(): void
    {
        $this->assertNull(SizeParser::parse(null));
        $this->assertNull(SizeParser::parse(''));
        $this->assertNull(SizeParser::parse('   '));
    }

    public function testParseInteger(): void
    {
        $this->assertSame(0, SizeParser::parse(0));
        $this->assertSame(1024, SizeParser::parse(1024));
        $this->assertSame(20971520, SizeParser::parse(20971520));
    }

    public function testParseNegativeIntegerThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Size cannot be negative');
        SizeParser::parse(-10);
    }

    public function testParseStringUnits(): void
    {
        $this->assertSame(500, SizeParser::parse('500'));
        $this->assertSame(500, SizeParser::parse('500B'));
        $this->assertSame(1024, SizeParser::parse('1K'));
        $this->assertSame(2048, SizeParser::parse('2KB'));
        $this->assertSame(20971520, SizeParser::parse('20M'));
        $this->assertSame(20971520, SizeParser::parse('20MB'));
        $this->assertSame(1572864, SizeParser::parse('1.5M'));
        $this->assertSame(1073741824, SizeParser::parse('1G'));
        $this->assertSame(1073741824, SizeParser::parse('1GB'));
        $this->assertSame(1024, SizeParser::parse('  1k  '));
    }

    public function testParseInvalidFormatThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Invalid size format');
        SizeParser::parse('abc');
    }

    public function testParseUnknownUnitThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Unknown size unit "TB"');
        SizeParser::parse('10TB');
    }
}
