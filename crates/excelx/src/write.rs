use rust_xlsxwriter::Workbook;

use crate::{CellValue, ExcelError, ExcelRow, validate_columns};

/// Convert a slice of row values into a single-sheet XLSX workbook.
pub fn to_xlsx<T: ExcelRow>(data: &[T]) -> Result<Vec<u8>, ExcelError> {
    let columns = T::columns();
    let sorted_columns = validate_columns(&columns)?;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for (write_col, column) in sorted_columns.iter().enumerate() {
        worksheet.write_string(0, as_col(write_col)?, column.header)?;
    }

    for (row_index, item) in data.iter().enumerate() {
        let values = item.to_row();
        if values.len() != columns.len() {
            return Err(ExcelError::Schema(format!(
                "to_row returned {} values but schema defines {} columns",
                values.len(),
                columns.len()
            )));
        }

        let excel_row = as_row(row_index + 1)?;
        for (write_col, column) in sorted_columns.iter().enumerate() {
            let schema_index = columns
                .iter()
                .position(|candidate| candidate.field == column.field)
                .ok_or_else(|| {
                    ExcelError::Schema(format!("schema column `{}` disappeared", column.field))
                })?;
            write_cell(
                worksheet,
                excel_row,
                as_col(write_col)?,
                &values[schema_index],
            )?;
        }
    }

    workbook.save_to_buffer().map_err(ExcelError::from)
}

fn write_cell(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: &CellValue,
) -> Result<(), ExcelError> {
    match value {
        CellValue::String(value) => {
            worksheet.write_string(row, col, value)?;
        }
        CellValue::Int(value) => {
            worksheet.write_number(row, col, *value as f64)?;
        }
        CellValue::Float(value) => {
            worksheet.write_number(row, col, *value)?;
        }
        CellValue::Bool(value) => {
            worksheet.write_boolean(row, col, *value)?;
        }
        CellValue::Empty => {}
    }

    Ok(())
}

fn as_row(value: usize) -> Result<u32, ExcelError> {
    u32::try_from(value)
        .map_err(|_| ExcelError::Write(format!("row index {value} exceeds XLSX limits")))
}

fn as_col(value: usize) -> Result<u16, ExcelError> {
    u16::try_from(value)
        .map_err(|_| ExcelError::Write(format!("column index {value} exceeds XLSX limits")))
}
