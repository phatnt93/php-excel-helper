<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Reader;

use ExcelHelper\Reader\ChunkReadFilter;
use ExcelHelper\Tests\TestCase;

final class ChunkReadFilterTest extends TestCase
{
    public function testReadCellBounds(): void
    {
        $filter = new ChunkReadFilter(
            startRow: 5,
            endRow: 10,
            headerRow: 1,
            allowedColumns: ['A', 'C'],
        );

        // Header row always read
        $this->assertTrue($filter->readCell('A', 1));
        $this->assertTrue($filter->readCell('B', 1));

        // Before start row
        $this->assertFalse($filter->readCell('A', 4));

        // After end row
        $this->assertFalse($filter->readCell('A', 11));

        // Between start and end row
        $this->assertTrue($filter->readCell('A', 5));
        $this->assertTrue($filter->readCell('C', 7));
        $this->assertFalse($filter->readCell('B', 7)); // B not in allowedColumns

        // Update rows
        $filter->setRows(20, 25);
        $this->assertFalse($filter->readCell('A', 5));
        $this->assertTrue($filter->readCell('A', 22));
    }
}
