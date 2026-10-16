<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Writer;

use ExcelHelper\Tests\TestCase;
use ExcelHelper\Writer\SheetNameSanitizer;

final class SheetNameSanitizerTest extends TestCase
{
    public function testSanitizesForbiddenCharacters(): void
    {
        $sanitized = SheetNameSanitizer::sanitize('Report*2026:Test?/[\\a]');
        $this->assertSame('Report_2026_Test____a_', $sanitized);
    }

    public function testTruncatesTo31Characters(): void
    {
        $long = 'This is a very long worksheet title exceeding thirty one characters';
        $sanitized = SheetNameSanitizer::sanitize($long);
        $this->assertLessThanOrEqual(31, mb_strlen($sanitized));
    }

    public function testEmptyBecomesSheet(): void
    {
        $this->assertSame('Sheet', SheetNameSanitizer::sanitize(''));
        $this->assertSame('Sheet', SheetNameSanitizer::sanitize('   '));
    }

    public function testDeduplicationAppendsSuffix(): void
    {
        $name1 = SheetNameSanitizer::sanitize('Users', ['Other']);
        $this->assertSame('Users', $name1);

        $name2 = SheetNameSanitizer::sanitize('Users', ['Users']);
        $this->assertSame('Users (2)', $name2);

        $name3 = SheetNameSanitizer::sanitize('Users', ['Users', 'users (2)']);
        $this->assertSame('Users (3)', $name3);
    }
}
