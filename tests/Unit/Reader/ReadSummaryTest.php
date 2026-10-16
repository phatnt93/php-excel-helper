<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Reader;

use ExcelHelper\Reader\ReadSummary;
use ExcelHelper\Reader\RowError;
use ExcelHelper\Tests\TestCase;

final class ReadSummaryTest extends TestCase
{
    public function testReadSummaryGetters(): void
    {
        $error = new RowError(1, 'A', 'code', 'required', null, 'Required');
        $summary = new ReadSummary(
            totalRows: 100,
            totalChunks: 2,
            errors: [$error],
            headers: ['A' => 'Code'],
            sheetName: 'Sheet1',
        );

        $this->assertSame(100, $summary->totalRows());
        $this->assertSame(2, $summary->totalChunks());
        $this->assertSame(1, $summary->errorCount());
        $this->assertTrue($summary->hasErrors());
        $this->assertSame([$error], $summary->errors());
        $this->assertSame(['A' => 'Code'], $summary->headers());
        $this->assertSame('Sheet1', $summary->sheetName());
    }
}
