//! Ergonomic single-sheet XLSX conversion for manually implemented row types.
//!
//! Phase 1 intentionally keeps the public API small. Implement [`ExcelRow`] for
//! your struct, then use [`to_xlsx`] and [`from_xlsx`] to round-trip values.

mod column;
mod error;
mod read;
mod row;
mod sheet;
mod types;
mod write;

pub use column::{ColumnDef, validate_columns};
pub use error::ExcelError;
pub use read::{from_reader, from_xlsx};
pub use row::{ExcelRow, RowView};
pub use sheet::SheetOptions;
pub use types::CellValue;
pub use write::to_xlsx;
