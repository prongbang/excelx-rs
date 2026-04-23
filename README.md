# excelx-rs

`excelx` is a small Rust crate for converting struct collections to a single
XLSX worksheet and parsing them back with explicit header and column-order
metadata.

Phase 1 supports manual `ExcelRow` implementations only. Derive macros,
defaults, and multi-sheet workbooks are planned for later phases.

## Example

```rust
use excelx::{CellValue, ColumnDef, ExcelError, ExcelRow, RowView, from_xlsx, to_xlsx};

#[derive(Debug, PartialEq)]
struct Person {
    id: i64,
    name: String,
    active: bool,
}

impl ExcelRow for Person {
    fn columns() -> Vec<ColumnDef> {
        vec![
            ColumnDef::new("id", "ID", 1),
            ColumnDef::new("name", "Name", 2),
            ColumnDef::new("active", "Active", 3),
        ]
    }

    fn to_row(&self) -> Vec<CellValue> {
        vec![
            self.id.into(),
            self.name.clone().into(),
            self.active.into(),
        ]
    }

    fn from_row(row: &RowView) -> Result<Self, ExcelError> {
        Ok(Self {
            id: row.required_i64("id")?,
            name: row.required_string("name")?,
            active: row.required_bool("active")?,
        })
    }
}

let people = vec![Person {
    id: 1,
    name: "Ada".to_owned(),
    active: true,
}];

let bytes = to_xlsx(&people)?;
let parsed = from_xlsx::<Person>(&bytes)?;
assert_eq!(parsed, people);
# Ok::<(), ExcelError>(())
```
