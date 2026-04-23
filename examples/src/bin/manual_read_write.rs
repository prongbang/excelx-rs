use std::fs;
use std::path::PathBuf;

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
        vec![self.id.into(), self.name.clone().into(), self.active.into()]
    }

    fn from_row(row: &RowView) -> Result<Self, ExcelError> {
        Ok(Self {
            id: row.required_i64("id")?,
            name: row.required_string("name")?,
            active: row.required_bool("active")?,
        })
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let people = vec![
        Person {
            id: 1,
            name: "Ada Lovelace".to_owned(),
            active: true,
        },
        Person {
            id: 2,
            name: "Grace Hopper".to_owned(),
            active: false,
        },
    ];

    let bytes = to_xlsx(&people)?;
    let parsed = from_xlsx::<Person>(&bytes)?;
    assert_eq!(parsed, people);

    let path = example_path("manual_read_write.xlsx")?;
    fs::write(&path, bytes)?;
    println!("wrote {}", path.display());

    Ok(())
}

fn example_path(file_name: &str) -> Result<PathBuf, std::io::Error> {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("../target/examples");
    fs::create_dir_all(&path)?;
    path.push(file_name);
    Ok(path)
}
