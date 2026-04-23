use std::collections::HashSet;

use crate::ExcelError;

/// Defines how a Rust field maps to a visible XLSX column.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ColumnDef {
    pub field: &'static str,
    pub header: &'static str,
    pub order: usize,
    pub default: Option<&'static str>,
}

impl ColumnDef {
    pub const fn new(field: &'static str, header: &'static str, order: usize) -> Self {
        Self {
            field,
            header,
            order,
            default: None,
        }
    }

    pub const fn with_default(
        field: &'static str,
        header: &'static str,
        order: usize,
        default: &'static str,
    ) -> Self {
        Self {
            field,
            header,
            order,
            default: Some(default),
        }
    }
}

/// Validate schema metadata and return columns sorted by ascending order.
pub fn validate_columns(columns: &[ColumnDef]) -> Result<Vec<ColumnDef>, ExcelError> {
    let mut orders = HashSet::with_capacity(columns.len());
    let mut headers = HashSet::with_capacity(columns.len());
    let mut fields = HashSet::with_capacity(columns.len());

    for column in columns {
        if column.field.trim().is_empty() {
            return Err(ExcelError::Schema(
                "column field cannot be empty".to_owned(),
            ));
        }

        if column.header.trim().is_empty() {
            return Err(ExcelError::Schema(format!(
                "column `{}` header cannot be empty",
                column.field
            )));
        }

        if !orders.insert(column.order) {
            return Err(ExcelError::DuplicateColumnOrder(column.order));
        }

        if !headers.insert(column.header) {
            return Err(ExcelError::DuplicateHeader(column.header.to_owned()));
        }

        if !fields.insert(column.field) {
            return Err(ExcelError::Schema(format!(
                "duplicate field mapping: {}",
                column.field
            )));
        }
    }

    let mut sorted = columns.to_vec();
    sorted.sort_by_key(|column| column.order);
    Ok(sorted)
}
