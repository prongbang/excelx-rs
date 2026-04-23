use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};

use calamine::{Data, Reader, Xlsx};

use crate::{CellValue, ColumnDef, ExcelError, ExcelRow, RowView, validate_columns};

/// Parse the first worksheet in an XLSX byte slice into typed rows.
pub fn from_xlsx<T: ExcelRow>(bytes: &[u8]) -> Result<Vec<T>, ExcelError> {
    from_reader(Cursor::new(bytes))
}

/// Parse the first worksheet in an XLSX reader into typed rows.
pub fn from_reader<T: ExcelRow, R: Read + Seek>(reader: R) -> Result<Vec<T>, ExcelError> {
    let columns = T::columns();
    let sorted_columns = validate_columns(&columns)?;
    let mut workbook: Xlsx<R> = Xlsx::new(reader)?;
    let range = workbook.worksheet_range_at(0).ok_or_else(|| {
        ExcelError::Parse("workbook does not contain any worksheets".to_owned())
    })??;

    let mut rows = range.rows();
    let header_row = rows
        .next()
        .ok_or_else(|| ExcelError::Parse("worksheet does not contain a header row".to_owned()))?;
    let header_map = build_header_map(header_row)?;
    ensure_required_headers(&sorted_columns, &header_map)?;

    let mut parsed_rows = Vec::new();
    for (relative_index, row) in rows.enumerate() {
        if is_empty_row(row) {
            continue;
        }

        let excel_row_number = relative_index + 2;
        let values = values_for_schema(row, &columns, &header_map)?;
        let row_view = RowView::new(excel_row_number, &columns, values);
        parsed_rows.push(T::from_row(&row_view)?);
    }

    Ok(parsed_rows)
}

fn build_header_map(row: &[Data]) -> Result<HashMap<String, usize>, ExcelError> {
    let mut headers = HashMap::with_capacity(row.len());

    for (index, cell) in row.iter().enumerate() {
        if matches!(cell, Data::Empty) {
            continue;
        }

        let header = match cell {
            Data::String(value) => value.trim().to_owned(),
            Data::Int(value) => value.to_string(),
            Data::Float(value) => value.to_string(),
            Data::Bool(value) => value.to_string(),
            other => {
                return Err(ExcelError::Parse(format!(
                    "unsupported header cell type at column {}: {other}",
                    index + 1
                )));
            }
        };

        if header.is_empty() {
            continue;
        }

        if headers.insert(header.clone(), index).is_some() {
            return Err(ExcelError::DuplicateHeader(header));
        }
    }

    Ok(headers)
}

fn ensure_required_headers(
    columns: &[ColumnDef],
    header_map: &HashMap<String, usize>,
) -> Result<(), ExcelError> {
    for column in columns {
        if !header_map.contains_key(column.header) {
            return Err(ExcelError::MissingHeader(column.header.to_owned()));
        }
    }

    Ok(())
}

fn values_for_schema(
    row: &[Data],
    columns: &[ColumnDef],
    header_map: &HashMap<String, usize>,
) -> Result<Vec<CellValue>, ExcelError> {
    columns
        .iter()
        .map(|column| {
            let index = header_map
                .get(column.header)
                .ok_or_else(|| ExcelError::MissingHeader(column.header.to_owned()))?;

            row.get(*index)
                .map(CellValue::try_from)
                .unwrap_or(Ok(CellValue::Empty))
        })
        .collect()
}

fn is_empty_row(row: &[Data]) -> bool {
    row.iter().all(|cell| matches!(cell, Data::Empty))
}
