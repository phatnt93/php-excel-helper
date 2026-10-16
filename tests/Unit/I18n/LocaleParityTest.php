<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\I18n;

use ExcelHelper\I18n\Translator;
use ExcelHelper\Tests\TestCase;

final class LocaleParityTest extends TestCase
{
    public function testAllLocalesHaveIdenticalKeysAndPlaceholders(): void
    {
        $enCatalog = Translator::loadCatalog('en');
        $viCatalog = Translator::loadCatalog('vi');

        $enKeys = array_keys($enCatalog);
        $viKeys = array_keys($viCatalog);
        sort($enKeys);
        sort($viKeys);

        $this->assertSame($enKeys, $viKeys, 'English and Vietnamese catalogs must have identical keys.');

        foreach ($enCatalog as $key => $enTemplate) {
            $viTemplate = $viCatalog[$key] ?? '';

            preg_match_all('/\{([a-zA-Z0-9_]+)\}/', $enTemplate, $enMatches);
            preg_match_all('/\{([a-zA-Z0-9_]+)\}/', $viTemplate, $viMatches);

            $enPlaceholders = $enMatches[1];
            $viPlaceholders = $viMatches[1];
            sort($enPlaceholders);
            sort($viPlaceholders);

            $this->assertSame(
                $enPlaceholders,
                $viPlaceholders,
                sprintf('Placeholders in key "%s" do not match between en and vi.', $key),
            );
        }
    }
}
