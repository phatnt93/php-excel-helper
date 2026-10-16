<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Writer;

use ExcelHelper\Excel;
use ExcelHelper\Tests\TestCase;

final class ExcelOutputTest extends TestCase
{
    public function testDownloadStreamsToOutputBuffer(): void
    {
        $out = Excel::write([
            'columns' => ['a' => 'Col A'],
            'data' => [['a' => 123]],
        ]);

        ob_start();
        $out->download('test.csv');
        $captured = ob_get_clean();

        $this->assertIsString($captured);
        $this->assertStringContainsString('Col A', $captured);
        $this->assertStringContainsString('123', $captured);
    }

    public function testSaveFormatInferenceFromExtension(): void
    {
        $csvPath = $this->tempPath('csv');
        $out = Excel::write([
            'columns' => ['a' => 'Header A'],
            'data' => [['a' => 'Val A']],
        ]);

        $saved = $out->save($csvPath);
        $this->assertFileExists($saved);
        $content = file_get_contents($saved);
        $this->assertIsString($content);
        $this->assertStringContainsString('Header A', $content);
    }
}
