use std::io::Cursor;

use calamine::{Reader, Xlsx};
use excelx::{
    CellValue, ColumnDef, ExcelError, ExcelRow, RowView, SheetData, from_xlsx, to_xlsx_multi,
};

#[derive(Debug, PartialEq)]
struct Account {
    id: i64,
    status: String,
    quota: i64,
    ratio: f64,
    active: bool,
}

impl ExcelRow for Account {
    fn columns() -> Vec<ColumnDef> {
        vec![
            ColumnDef::new("id", "ID", 1),
            ColumnDef::with_default("status", "Status", 2, "new"),
            ColumnDef::with_default("quota", "Quota", 3, "10"),
            ColumnDef::with_default("ratio", "Ratio", 4, "0.75"),
            ColumnDef::with_default("active", "Active", 5, "true"),
        ]
    }

    fn to_row(&self) -> Vec<CellValue> {
        vec![
            self.id.into(),
            CellValue::Empty,
            CellValue::Empty,
            CellValue::Empty,
            CellValue::Empty,
        ]
    }

    fn from_row(row: &RowView) -> Result<Self, ExcelError> {
        Ok(Self {
            id: row.required_i64("id")?,
            status: row.required_string("status")?,
            quota: row.required_i64("quota")?,
            ratio: row.required_f64("ratio")?,
            active: row.required_bool("active")?,
        })
    }
}

#[test]
fn applies_defaults_when_cells_are_empty() {
    let bytes = excelx::to_xlsx(&[Account {
        id: 1,
        status: "ignored".to_owned(),
        quota: 0,
        ratio: 0.0,
        active: false,
    }])
    .expect("write workbook");

    let parsed = from_xlsx::<Account>(&bytes).expect("parse workbook");

    assert_eq!(
        parsed,
        vec![Account {
            id: 1,
            status: "new".to_owned(),
            quota: 10,
            ratio: 0.75,
            active: true,
        }]
    );
}

#[test]
fn reports_invalid_default_with_schema_context() {
    #[derive(Debug)]
    struct InvalidDefault;

    impl ExcelRow for InvalidDefault {
        fn columns() -> Vec<ColumnDef> {
            vec![
                ColumnDef::new("id", "ID", 1),
                ColumnDef::with_default("quota", "Quota", 2, "many"),
            ]
        }

        fn to_row(&self) -> Vec<CellValue> {
            vec![CellValue::Int(1), CellValue::Empty]
        }

        fn from_row(row: &RowView) -> Result<Self, ExcelError> {
            row.required_i64("quota")?;
            Ok(Self)
        }
    }

    let bytes = excelx::to_xlsx(&[InvalidDefault]).expect("write workbook");
    let error = from_xlsx::<InvalidDefault>(&bytes).expect_err("invalid default");

    assert!(matches!(
        error,
        ExcelError::InvalidDefault {
            field,
            header,
            expected,
            value,
        } if field == "quota" && header == "Quota" && expected == "integer" && value == "many"
    ));
}

#[test]
fn writes_multiple_sheets() {
    let sheets = vec![
        SheetData::new(
            "Accounts",
            vec![Account {
                id: 1,
                status: "ignored".to_owned(),
                quota: 0,
                ratio: 0.0,
                active: false,
            }],
        ),
        SheetData::new(
            "Archive",
            vec![Account {
                id: 2,
                status: "ignored".to_owned(),
                quota: 0,
                ratio: 0.0,
                active: false,
            }],
        ),
    ];

    let bytes = to_xlsx_multi(&sheets).expect("write multi-sheet workbook");
    let workbook: Xlsx<_> = Xlsx::new(Cursor::new(bytes)).expect("open workbook");

    assert_eq!(workbook.sheet_names(), ["Accounts", "Archive"]);
}

#[test]
fn rejects_duplicate_sheet_names() {
    let sheets = vec![
        SheetData::new("Accounts", Vec::<Account>::new()),
        SheetData::new("accounts", Vec::<Account>::new()),
    ];

    let error = to_xlsx_multi(&sheets).expect_err("duplicate sheet names");

    assert!(
        matches!(error, ExcelError::Schema(message) if message.contains("duplicate sheet name"))
    );
}
