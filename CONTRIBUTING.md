# Contributing to php-excel-helper

Thank you for considering contributing to `phatnt93/php-excel-helper`! We welcome bug fixes, documentation improvements, new features, and translations.

---

## Development Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/phatnt93/php-excel-helper.git
   cd php-excel-helper
   ```

2. **Install dependencies:**
   ```bash
   composer install
   ```

3. **Verify your local environment:**
   - PHP: 8.1 or higher
   - Extensions required by PhpSpreadsheet: `ctype`, `dom`, `fileinfo`, `gd`, `iconv`, `libxml`, `mbstring`, `simplexml`, `xml`, `xmlreader`, `xmlwriter`, `zip`, `zlib`

---

## Quality & Testing Workflow

Before submitting a pull request, ensure all checks pass:

```bash
composer validate --strict
composer cs:check
composer analyse
composer test
composer test:coverage
for f in examples/*.php; do php "$f"; done
```

---

## Coding Standards & Conventions

- **Strict Types:** Every PHP file must declare `declare(strict_types=1);` as the first statement.
- **Namespaces:** PSR-4 under `ExcelHelper\`. Interfaces belong in `ExcelHelper\Contracts\`.
- **Array config keys** use `snake_case`; matching fluent builder methods use `camelCase`. Every option must exist in both and be covered by a parity test.
- **Exception Messages:** developer-facing exception messages are in English and built with named constructors.
- **PhpSpreadsheet compatibility:** only use APIs available in `phpoffice/phpspreadsheet` 5.8.0 (CI runs with lowest dependencies).
- **Security First:**
  - Never trust the file extension; detect the format by content.
  - Never write user strings as formulas.
  - Every regex using `$` must include the `D` flag or use `\z`.

---

## Adding a New Translation (Locale)

1. Copy `src/I18n/en.php` to `src/I18n/{locale}.php`.
2. Translate all message strings while preserving exact `{placeholder}` keys.
3. Add the locale code to `tests/Unit/I18n/LocaleParityTest.php`.
4. Run `composer test` to confirm key and placeholder parity.
5. List the new locale in `README.md` and `llms-full.txt`.

---

## Commit Message Convention

We follow [Conventional Commits](https://www.conventionalcommits.org/): `feat:`, `fix:`, `test:`, `docs:`, `refactor:`, `style:`, `chore:`.

---

## Reporting Vulnerabilities

Please **do not** open public GitHub issues for security vulnerabilities. See [SECURITY.md](SECURITY.md).
