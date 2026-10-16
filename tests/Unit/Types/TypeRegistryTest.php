<?php

declare(strict_types=1);

namespace ExcelHelper\Tests\Unit\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\InvalidConfigException;
use ExcelHelper\Tests\TestCase;
use ExcelHelper\Types\CellValue;
use ExcelHelper\Types\ColumnContext;
use ExcelHelper\Types\TypeRegistry;

final class TypeRegistryTest extends TestCase
{
    public function testGetBuiltInTypes(): void
    {
        $types = [
            'string', 'int', 'integer', 'float', 'number', 'decimal', 'bool', 'boolean',
            'date', 'datetime', 'time', 'raw', 'formatted',
            'auto', 'percent', 'money', 'formula', 'link',
        ];

        foreach ($types as $type) {
            $this->assertTrue(TypeRegistry::has($type), "TypeRegistry should have '{$type}'");
            $instance = TypeRegistry::get($type);
            $this->assertInstanceOf(TypeInterface::class, $instance);
        }
    }

    public function testGetCaseInsensitive(): void
    {
        $this->assertInstanceOf(TypeInterface::class, TypeRegistry::get('STRING'));
        $this->assertInstanceOf(TypeInterface::class, TypeRegistry::get('  Int  '));
    }

    public function testGetUnknownTypeThrows(): void
    {
        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Unknown column type "unknown_custom_type"');
        TypeRegistry::get('unknown_custom_type');
    }

    public function testRegisterCustomType(): void
    {
        $custom = new class implements TypeInterface {
            public function name(): string
            {
                return 'custom_badge';
            }

            public function read(mixed $value, ColumnContext $context): mixed
            {
                return 'badge_' . $value;
            }

            public function write(mixed $value, ColumnContext $context): CellValue
            {
                return CellValue::string((string) $value);
            }
        };

        TypeRegistry::register($custom);
        $this->assertTrue(TypeRegistry::has('custom_badge'));
        $this->assertSame($custom, TypeRegistry::get('custom_badge'));
    }

    public function testRegisterBuiltInWithoutReplaceThrows(): void
    {
        $custom = new class implements TypeInterface {
            public function name(): string
            {
                return 'string';
            }

            public function read(mixed $value, ColumnContext $context): mixed
            {
                return $value;
            }

            public function write(mixed $value, ColumnContext $context): CellValue
            {
                return CellValue::string((string) $value);
            }
        };

        $this->expectException(InvalidConfigException::class);
        $this->expectExceptionMessage('Cannot overwrite built-in type "string" without setting $replace = true.');
        TypeRegistry::register($custom);
    }

    public function testRegisterBuiltInWithReplaceSucceeds(): void
    {
        $custom = new class implements TypeInterface {
            public function name(): string
            {
                return 'string';
            }

            public function read(mixed $value, ColumnContext $context): mixed
            {
                return 'replaced';
            }

            public function write(mixed $value, ColumnContext $context): CellValue
            {
                return CellValue::string('replaced');
            }
        };

        TypeRegistry::register($custom, replace: true);
        $this->assertSame($custom, TypeRegistry::get('string'));

        TypeRegistry::reset();
        $this->assertNotSame($custom, TypeRegistry::get('string'));
    }
}
