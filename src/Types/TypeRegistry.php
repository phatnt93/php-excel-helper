<?php

declare(strict_types=1);

namespace ExcelHelper\Types;

use ExcelHelper\Contracts\TypeInterface;
use ExcelHelper\Exceptions\InvalidConfigException;

/**
 * Registry of available column types.
 */
final class TypeRegistry
{
    /**
     * @var array<string, TypeInterface>
     */
    private static array $types = [];

    /**
     * @var array<string, bool>
     */
    private static array $builtIns = [];

    public static function get(string $name): TypeInterface
    {
        self::ensureInitialized();

        $key = strtolower(trim($name));

        if (!isset(self::$types[$key])) {
            $available = implode(', ', self::availableTypes());

            throw InvalidConfigException::withMessage(sprintf(
                'Unknown column type "%s". Available types are: %s.',
                $name,
                $available,
            ));
        }

        return self::$types[$key];
    }

    public static function has(string $name): bool
    {
        self::ensureInitialized();

        return isset(self::$types[strtolower(trim($name))]);
    }

    public static function register(TypeInterface $type, bool $replace = false): void
    {
        self::ensureInitialized();

        $name = strtolower(trim($type->name()));

        if (isset(self::$builtIns[$name]) && !$replace) {
            throw InvalidConfigException::withMessage(sprintf(
                'Cannot overwrite built-in type "%s" without setting $replace = true.',
                $type->name(),
            ));
        }

        self::$types[$name] = $type;
    }

    public static function reset(): void
    {
        self::$types = [];
        self::$builtIns = [];
        self::initialize();
    }

    /**
     * @return list<string>
     */
    public static function availableTypes(): array
    {
        self::ensureInitialized();

        $keys = array_keys(self::$types);
        sort($keys);

        return $keys;
    }

    private static function ensureInitialized(): void
    {
        if (self::$types === []) {
            self::initialize();
        }
    }

    private static function initialize(): void
    {
        $builtInInstances = [
            new StringType(),
            new IntType(),
            new FloatType(),
            new DecimalType(),
            new BoolType(),
            new DateType(),
            new DateTimeType(),
            new TimeType(),
            new RawType(),
            new FormattedType(),
            new AutoType(),
            new PercentType(),
            new MoneyType(),
            new FormulaType(),
            new LinkType(),
        ];

        foreach ($builtInInstances as $instance) {
            $name = strtolower($instance->name());
            self::$types[$name] = $instance;
            self::$builtIns[$name] = true;
        }

        // Type aliases
        self::$types['number'] = self::$types['float'];
        self::$builtIns['number'] = true;

        self::$types['integer'] = self::$types['int'];
        self::$builtIns['integer'] = true;

        self::$types['boolean'] = self::$types['bool'];
        self::$builtIns['boolean'] = true;
    }
}
