//! Ergonomic XLSX conversion for manually implemented row types.
//!
//! Implement [`ExcelRow`] for your struct, then use [`to_xlsx`] and
//! [`from_xlsx`] to round-trip a first worksheet, or the sheet-selection and
//! multi-sheet APIs for workbooks with multiple homogeneous worksheets.

mod column;
mod error;
mod read;
mod row;
mod sheet;
mod types;
mod write;

pub use column::{ColumnDef, validate_columns};
pub use error::ExcelError;
pub use read::{
    ParsedSheet, ReadOptions, SheetRef, from_reader, from_reader_multi, from_reader_with_options,
    from_xlsx, from_xlsx_multi, from_xlsx_sheet, from_xlsx_with_options,
};
pub use row::{ExcelRow, RowView};
pub use sheet::{SheetData, SheetOptions};
pub use types::CellValue;
pub use write::{to_xlsx, to_xlsx_multi};
