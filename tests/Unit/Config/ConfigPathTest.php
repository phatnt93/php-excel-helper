<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Config;

use ExcelHelper\Config\ConfigPath;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Tests\TestCase;

final class ConfigPathTest extends TestCase
{
    public function testEmptyRootPath(): void
    {
        $path = ConfigPath::root();
        $this->assertTrue($path->isEmpty());
        $this->assertSame('', $path->toString());
        $this->assertSame('', (string) $path);
        $this->assertSame([], $path->segments());
    }

    public function testPathWithSegments(): void
    {
        $path = ConfigPath::root('sheets')
            ->append(0)
            ->append('columns')
            ->append('salary')
            ->append('type');

        $this->assertFalse($path->isEmpty());
        $this->assertSame('sheets.0.columns.salary.type', $path->toString());
        $this->assertSame('sheets.0.columns.salary.type', (string) $path);
        $this->assertSame(['sheets', 0, 'columns', 'salary', 'type'], $path->segments());
    }

    public function testAssertKnownKeysPassesOnValidConfig(): void
    {
        $config = ['sheet' => 0, 'start_row' => 2];
        $path = ConfigPath::root('options');

        ConfigPath::assertKnownKeys($config, ['sheet', 'start_row', 'end_row'], $path);
        $this->addToAssertionCount(1);
    }

    public function testAssertKnownKeysThrowsOnUnknownKey(): void
    {
        $config = ['strat_row' => 2];
        $path = ConfigPath::root('options');

        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Configuration error at "options.strat_row": Unknown configuration option "strat_row"');

        ConfigPath::assertKnownKeys($config, ['start_row', 'end_row'], $path);
    }
}
