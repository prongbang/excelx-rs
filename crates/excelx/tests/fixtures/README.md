# Compatibility Fixtures

`compatibility_workbooks.rs` reads `.xlsx` files from the directory in
`EXCELX_COMPAT_FIXTURE_DIR`.

The CI workflow generates a LibreOffice workbook at runtime with
`libreoffice --headless --convert-to xlsx`. Excel-created workbooks should be
added here manually when available, using the same visible sheet shape:

```text
ID,Name,Active,Score
1,Ada,true,98.5
2,Grace,false,88
```

Local run example:

```sh
EXCELX_COMPAT_FIXTURE_DIR=crates/excelx/tests/fixtures/compat cargo test -p excelx --test compatibility_workbooks
```
