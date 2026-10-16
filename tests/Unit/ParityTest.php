<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit;

use ExcelHelper\Config\WriteConfigParser;
use ExcelHelper\Excel;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Writer\SheetBuilder;

final class ParityTest extends TestCase
{
    public function testReaderParity(): void
    {
        $path = $this->createSpreadsheetFile([
            ['Mã NV', 'Họ Tên', 'Lương'],
            ['NV1', 'An', 1000],
            ['NV2', 'Bình', 2000],
        ]);

        $arrayOptions = [
            'header_row' => 1,
            'start_row' => 2,
            'end_row' => 3,
            'on_error' => 'skip_row',
            'columns' => [
                'A' => 'code',
                'B' => 'name',
                'C' => ['key' => 'salary', 'type' => 'int'],
            ],
        ];

        $arrayResult = Excel::read($path, $arrayOptions);

        $builderResult = Excel::reader($path)
            ->headerRow(1)
            ->startRow(2)
            ->endRow(3)
            ->onError('skip_row')
            ->column('A', 'code')
            ->column('B', 'name')
            ->column('C', 'salary', type: 'int')
            ->read();

        $this->assertSame($arrayResult->toArray(), $builderResult->toArray());
    }

    public function testWriterParity(): void
    {
        $data = [
            ['id' => '1', 'name' => 'Alice'],
            ['id' => '2', 'name' => 'Bob'],
        ];

        $arrayConfig = [
            'format' => 'csv',
            'columns' => [
                'id' => 'ID',
                'name' => 'Name',
            ],
            'data' => $data,
        ];

        $arrayOutput = Excel::write($arrayConfig)->toString();

        $builderOutput = Excel::writer()
            ->format('csv')
            ->column('id', 'ID')
            ->column('name', 'Name')
            ->data($data)
            ->toString();

        $this->assertSame($arrayOutput, $builderOutput);
    }

    public function testMultiSheetWriterParity(): void
    {
        $arrayConfig = [
            'sheets' => [
                [
                    'name' => 'S1',
                    'columns' => ['a' => 'A'],
                    'data' => [['a' => 1]],
                ],
                [
                    'name' => 'S2',
                    'columns' => ['b' => 'B'],
                    'data' => [['b' => 2]],
                ],
            ],
        ];

        $builder = Excel::writer()
            ->sheet('S1', function (SheetBuilder $s): void {
                $s->column('a', 'A')->data([['a' => 1]]);
            })
            ->sheet('S2', function (SheetBuilder $s): void {
                $s->column('b', 'B')->data([['b' => 2]]);
            });

        $defArray = WriteConfigParser::parse($arrayConfig);
        $defBuilder = WriteConfigParser::parse($builder->toConfig());

        $this->assertSame(count($defArray->sheets), count($defBuilder->sheets));
        $this->assertSame($defArray->sheets[0]->name, $defBuilder->sheets[0]->name);
        $this->assertSame($defArray->sheets[1]->name, $defBuilder->sheets[1]->name);
    }
}
