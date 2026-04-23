use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};

use calamine::{Data, Range, Reader, Xlsx};

use crate::{CellValue, ColumnDef, ExcelError, ExcelRow, RowView, validate_columns};

/// Selects a worksheet by workbook order or visible worksheet name.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SheetRef<'a> {
    Index(usize),
    Name(&'a str),
}

/// Options for parsing an XLSX workbook.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ReadOptions<'a> {
    pub sheet: SheetRef<'a>,
}

impl<'a> Default for ReadOptions<'a> {
    fn default() -> Self {
        Self {
            sheet: SheetRef::Index(0),
        }
    }
}

/// Parsed rows for one worksheet in a homogeneous multi-sheet workbook.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ParsedSheet<T> {
    pub name: String,
    pub rows: Vec<T>,
}

impl<T> ParsedSheet<T> {
    pub fn new(name: impl Into<String>, rows: Vec<T>) -> Self {
        Self {
            name: name.into(),
            rows,
        }
    }
}

/// Parse the first worksheet in an XLSX byte slice into typed rows.
pub fn from_xlsx<T: ExcelRow>(bytes: &[u8]) -> Result<Vec<T>, ExcelError> {
    from_reader(Cursor::new(bytes))
}

/// Parse the first worksheet in an XLSX reader into typed rows.
pub fn from_reader<T: ExcelRow, R: Read + Seek>(reader: R) -> Result<Vec<T>, ExcelError> {
    from_reader_with_options(reader, ReadOptions::default())
}

/// Parse one selected worksheet in an XLSX byte slice into typed rows.
pub fn from_xlsx_sheet<T: ExcelRow>(
    bytes: &[u8],
    sheet: SheetRef<'_>,
) -> Result<Vec<T>, ExcelError> {
    from_xlsx_with_options(bytes, ReadOptions { sheet })
}

/// Parse one selected worksheet in an XLSX byte slice with explicit options.
pub fn from_xlsx_with_options<T: ExcelRow>(
    bytes: &[u8],
    options: ReadOptions<'_>,
) -> Result<Vec<T>, ExcelError> {
    from_reader_with_options(Cursor::new(bytes), options)
}

/// Parse one selected worksheet in an XLSX reader with explicit options.
pub fn from_reader_with_options<T: ExcelRow, R: Read + Seek>(
    reader: R,
    options: ReadOptions<'_>,
) -> Result<Vec<T>, ExcelError> {
    let columns = T::columns();
    let sorted_columns = validate_columns(&columns)?;
    let mut workbook: Xlsx<R> = Xlsx::new(reader)?;
    let range = worksheet_range(&mut workbook, options.sheet)?;

    parse_range(range, &columns, &sorted_columns)
}

/// Parse all worksheets in an XLSX byte slice as the same row type.
pub fn from_xlsx_multi<T: ExcelRow>(bytes: &[u8]) -> Result<Vec<ParsedSheet<T>>, ExcelError> {
    from_reader_multi(Cursor::new(bytes))
}

/// Parse all worksheets in an XLSX reader as the same row type.
pub fn from_reader_multi<T: ExcelRow, R: Read + Seek>(
    reader: R,
) -> Result<Vec<ParsedSheet<T>>, ExcelError> {
    let columns = T::columns();
    let sorted_columns = validate_columns(&columns)?;
    let mut workbook: Xlsx<R> = Xlsx::new(reader)?;
    let sheet_names = workbook.sheet_names();

    if sheet_names.is_empty() {
        return Err(ExcelError::Parse(
            "workbook does not contain any worksheets".to_owned(),
        ));
    }

    let mut parsed_sheets = Vec::with_capacity(sheet_names.len());
    for name in sheet_names {
        let range = workbook.worksheet_range(&name)?;
        let rows = parse_range(range, &columns, &sorted_columns)?;
        parsed_sheets.push(ParsedSheet::new(name, rows));
    }

    Ok(parsed_sheets)
}

fn parse_range<T: ExcelRow>(
    range: Range<Data>,
    columns: &[ColumnDef],
    sorted_columns: &[ColumnDef],
) -> Result<Vec<T>, ExcelError> {
    let mut rows = range.rows();
    let header_row = rows
        .next()
        .ok_or_else(|| ExcelError::Parse("worksheet does not contain a header row".to_owned()))?;
    let header_map = build_header_map(header_row)?;
    ensure_required_headers(sorted_columns, &header_map)?;

    let mut parsed_rows = Vec::new();
    for (relative_index, row) in rows.enumerate() {
        if is_empty_row(row) {
            continue;
        }

        let excel_row_number = relative_index + 2;
        let values = values_for_schema(row, columns, &header_map)?;
        let row_view = RowView::new(excel_row_number, columns, values);
        parsed_rows.push(T::from_row(&row_view)?);
    }

    Ok(parsed_rows)
}

fn worksheet_range<R: Read + Seek>(
    workbook: &mut Xlsx<R>,
    sheet: SheetRef<'_>,
) -> Result<Range<Data>, ExcelError> {
    match sheet {
        SheetRef::Index(index) => workbook
            .worksheet_range_at(index)
            .ok_or_else(|| missing_sheet_by_index(index))?
            .map_err(ExcelError::from),
        SheetRef::Name(name) => {
            if !workbook
                .sheet_names()
                .iter()
                .any(|sheet_name| sheet_name == name)
            {
                return Err(missing_sheet_by_name(name));
            }

            workbook.worksheet_range(name).map_err(ExcelError::from)
        }
    }
}

fn missing_sheet_by_index(index: usize) -> ExcelError {
    if index == 0 {
        ExcelError::Parse("workbook does not contain any worksheets".to_owned())
    } else {
        ExcelError::Parse(format!("worksheet index {index} does not exist"))
    }
}

fn missing_sheet_by_name(name: &str) -> ExcelError {
    ExcelError::Parse(format!("worksheet named `{name}` does not exist"))
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
