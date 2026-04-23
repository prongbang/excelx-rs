use std::path::Path;

use excelx::{CellValue, ColumnDef, ExcelError, ExcelRow, RowView, from_xlsx};

#[derive(Debug, PartialEq)]
struct CompatPerson {
    id: i64,
    name: String,
    active: bool,
    score: f64,
}

impl ExcelRow for CompatPerson {
    fn columns() -> Vec<ColumnDef> {
        vec![
            ColumnDef::new("id", "ID", 1),
            ColumnDef::new("name", "Name", 2),
            ColumnDef::new("active", "Active", 3),
            ColumnDef::new("score", "Score", 4),
        ]
    }

    fn to_row(&self) -> Vec<CellValue> {
        vec![
            self.id.into(),
            self.name.clone().into(),
            self.active.into(),
            self.score.into(),
        ]
    }

    fn from_row(row: &RowView) -> Result<Self, ExcelError> {
        Ok(Self {
            id: row.required_i64("id")?,
            name: row.required_string("name")?,
            active: row.required_bool("active")?,
            score: row.required_f64("score")?,
        })
    }
}

#[test]
fn reads_external_compatibility_workbooks() {
    let Some(fixture_dir) = std::env::var_os("EXCELX_COMPAT_FIXTURE_DIR") else {
        return;
    };

    let fixture_dir = Path::new(&fixture_dir);
    let mut checked = 0usize;
    for entry in std::fs::read_dir(fixture_dir).expect("read compatibility fixture directory") {
        let path = entry.expect("read compatibility fixture entry").path();
        if path.extension().and_then(|extension| extension.to_str()) != Some("xlsx") {
            continue;
        }

        let bytes = std::fs::read(&path).expect("read compatibility workbook");
        let parsed = from_xlsx::<CompatPerson>(&bytes).unwrap_or_else(|error| {
            panic!("parse {}: {error}", path.display());
        });

        assert_eq!(
            parsed,
            vec![
                CompatPerson {
                    id: 1,
                    name: "Ada".to_owned(),
                    active: true,
                    score: 98.5,
                },
                CompatPerson {
                    id: 2,
                    name: "Grace".to_owned(),
                    active: false,
                    score: 88.0,
                },
            ],
            "unexpected rows in {}",
            path.display()
        );
        checked += 1;
    }

    assert!(
        checked > 0,
        "EXCELX_COMPAT_FIXTURE_DIR must contain at least one .xlsx fixture"
    );
}
