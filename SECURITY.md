# Security Policy

`phatnt93/php-excel-helper` reads spreadsheet files that are often uploaded by untrusted users and writes exports that may contain user-supplied data. We treat potential vulnerabilities (malicious file parsing, resource exhaustion, formula/CSV injection, header injection) seriously.

---

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.x     | :white_check_mark: |
| < 1.0   | :x:                |

Spreadsheet parsing is performed by `phpoffice/phpspreadsheet`. Keep it up to date; vulnerabilities in PhpSpreadsheet itself should be reported to the [PhpSpreadsheet project](https://github.com/PHPOffice/PhpSpreadsheet/security).

---

## Reporting a Vulnerability

**Please DO NOT report security vulnerabilities via public GitHub issues, discussions, or pull requests.**

1. **GitHub Private Vulnerability Reporting (Preferred):** open the [Security Advisories tab](https://github.com/phatnt93/php-excel-helper/security/advisories) and click **"Report a vulnerability"**.
2. **Email:** `phatnt93@gmail.com` (CC: `thanhphat.uit@gmail.com`) with the subject `[SECURITY] php-excel-helper vulnerability report`.

Please include a description and impact, reproduction steps or a minimal PoC (sample file if relevant), the library/PhpSpreadsheet/PHP versions tested, and any suggested mitigation.

---

## Response & Disclosure Process

1. **Initial Response:** within **48 hours**.
2. **Investigation & Triage:** verify the issue and assess severity.
3. **Patch & Release:** fix in a private fork, verified against the CI matrix, released as a patch version.
4. **Public Advisory:** GitHub Security Advisory (and CVE when applicable) published with the release, crediting the reporter unless anonymity is requested.
