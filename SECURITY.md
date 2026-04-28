# Security Policy

## Reporting a Vulnerability

We take security seriously. If you discover a vulnerability, please report it
responsibly and avoid public disclosure until a fix is available.

### Preferred: GitHub Security Advisory

1. Go to the [Security Advisories page](https://github.com/yeongseon/excel-dbapi/security/advisories/new)
2. Click "Report a vulnerability"
3. Submit the details in the private advisory

### Alternative: Email

If you prefer email, contact: yeongseon.choe@gmail.com

### What to Include

* A clear description of the issue
* Steps to reproduce
* Impact assessment
* Suggested mitigation or fix, if available

### Response Timeline

* Initial response: within 48 hours
* Status update: within 7 days

## Supported Versions

| Version | Supported |
| ------- | --------- |
| Latest release | :white_check_mark: |
| Older releases | :x: |

## Spreadsheet Formula Injection

### What is it?

Spreadsheet formula injection (also called CSV injection) occurs when untrusted
data written to a spreadsheet cell starts with characters that spreadsheet
applications interpret as formulas: `=`, `+`, `-`, `@`, `\t`, `\r`.

A malicious value like `=CMD("calc")` or `=HYPERLINK("https://evil.example",
"Click here")` could execute commands or exfiltrate data when the file is
opened in Excel or Google Sheets.

Reference: [OWASP — CSV Injection](https://owasp.org/www-community/attacks/CSV_Injection)

### How excel-dbapi defends against it

**Default behavior (`sanitize_formulas=True`)**:

All string values written via `INSERT` or `UPDATE` are checked before being
stored. If a value starts with any of the dangerous prefix characters
(`=`, `+`, `-`, `@`, `\t`, `\r`), it is prefixed with a single-quote (`'`)
so the spreadsheet treats it as a text literal rather than a formula.

This is enabled by default on all three engines (openpyxl, pandas, graph).

**Opting out**:

If you intentionally write formulas (e.g. `=SUM(A1:A10)`), pass
`sanitize_formulas=False` to `connect()`:

```python
from excel_dbapi import connect

conn = connect("workbook.xlsx", sanitize_formulas=False)
```

> **Warning**: Only disable sanitization when you fully control the data being
> written. Never disable it when writing user-supplied or external data.

### Scope

- **Covered**: All DML writes (`INSERT`, `UPDATE`) across all engines.
- **Not covered**: Data written via direct workbook access (`conn.workbook`)
  bypasses the SQL layer and its sanitization. If you use `.workbook`,
  you are responsible for your own sanitization.
