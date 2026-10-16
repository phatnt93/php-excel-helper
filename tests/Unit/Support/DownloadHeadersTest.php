<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Support;

use ExcelHelper\Support\DownloadHeaders;
use ExcelHelper\Tests\TestCase;

final class DownloadHeadersTest extends TestCase
{
    public function testCreateHeadersForXlsx(): void
    {
        $headers = DownloadHeaders::create('báo-cáo.xlsx', 'xlsx');

        $this->assertArrayHasKey('Content-Type', $headers);
        $this->assertSame('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', $headers['Content-Type']);
        $this->assertStringContainsString('b_o-c_o.xlsx', $headers['Content-Disposition']);
        $this->assertStringContainsString('filename*=UTF-8\'\'', $headers['Content-Disposition']);
    }

    public function testCreateHeadersStripsCrlf(): void
    {
        $headers = DownloadHeaders::create("evil\r\ninjection.xlsx", 'xlsx');

        $this->assertStringNotContainsString("\r", $headers['Content-Disposition']);
        $this->assertStringNotContainsString("\n", $headers['Content-Disposition']);
    }

    public function testCreateHeadersForCsv(): void
    {
        $headers = DownloadHeaders::create('export.csv', 'csv');

        $this->assertSame('text/csv; charset=UTF-8', $headers['Content-Type']);
    }
}
