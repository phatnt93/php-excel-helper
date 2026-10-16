<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Reader;

use ExcelHelper\Reader\ReadResult;
use ExcelHelper\Reader\RowError;
use ExcelHelper\Tests\TestCase;

final class ReadResultTest extends TestCase
{
    public function testReadResultGettersAndToArray(): void
    {
        $rows = [
            ['id' => 1, 'name' => 'Alice'],
            ['id' => 2, 'name' => 'Bob'],
        ];

        $error1 = new RowError(2, 'B', 'name', 'required', null, 'Required');
        $error2 = new RowError(2, 'A', 'id', 'invalid_int', 'foo', 'Invalid int');
        $error3 = new RowError(3, 'B', 'name', 'required', null, 'Required');

        $result = new ReadResult(
            rows: $rows,
            errors: [$error1, $error2, $error3],
            headers: ['A' => 'ID', 'B' => 'Name'],
            sheetName: 'Users',
        );

        $this->assertSame($rows, $result->rows());
        $this->assertCount(3, $result->errors());
        $this->assertTrue($result->hasErrors());
        $this->assertSame(['A' => 'ID', 'B' => 'Name'], $result->headers());
        $this->assertSame('Users', $result->sheetName());
        $this->assertSame(2, $result->rowCount());

        $errorsByRow = $result->errorsByRow();
        $this->assertCount(2, $errorsByRow[2]);
        $this->assertCount(1, $errorsByRow[3]);

        $arr = $result->toArray();
        $this->assertSame(2, $arr['rowCount']);
        $this->assertSame('Users', $arr['sheetName']);
        $this->assertCount(3, $arr['errors']);
    }
}
