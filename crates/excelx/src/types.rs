use calamine::Data;

use crate::ExcelError;

/// A small, explicit value model for the cell types supported in Phase 1.
#[derive(Clone, Debug, PartialEq)]
pub enum CellValue {
    String(String),
    Int(i64),
    Float(f64),
    Bool(bool),
    Empty,
}

impl CellValue {
    pub fn type_name(&self) -> &'static str {
        match self {
            Self::String(_) => "string",
            Self::Int(_) => "integer",
            Self::Float(_) => "float",
            Self::Bool(_) => "boolean",
            Self::Empty => "empty",
        }
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }
}

impl From<&str> for CellValue {
    fn from(value: &str) -> Self {
        Self::String(value.to_owned())
    }
}

impl From<String> for CellValue {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<i64> for CellValue {
    fn from(value: i64) -> Self {
        Self::Int(value)
    }
}

impl From<i32> for CellValue {
    fn from(value: i32) -> Self {
        Self::Int(i64::from(value))
    }
}

impl From<u32> for CellValue {
    fn from(value: u32) -> Self {
        Self::Int(i64::from(value))
    }
}

impl From<f64> for CellValue {
    fn from(value: f64) -> Self {
        Self::Float(value)
    }
}

impl From<f32> for CellValue {
    fn from(value: f32) -> Self {
        Self::Float(f64::from(value))
    }
}

impl From<bool> for CellValue {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl TryFrom<&Data> for CellValue {
    type Error = ExcelError;

    fn try_from(value: &Data) -> Result<Self, Self::Error> {
        match value {
            Data::String(value) => Ok(Self::String(value.clone())),
            Data::Int(value) => Ok(Self::Int(*value)),
            Data::Float(value) => Ok(Self::Float(*value)),
            Data::Bool(value) => Ok(Self::Bool(*value)),
            Data::Empty => Ok(Self::Empty),
            Data::DateTime(value) => Err(ExcelError::Parse(format!(
                "unsupported date/time cell value: {value}"
            ))),
            Data::DateTimeIso(value) => Err(ExcelError::Parse(format!(
                "unsupported ISO date/time cell value: {value}"
            ))),
            Data::DurationIso(value) => Err(ExcelError::Parse(format!(
                "unsupported ISO duration cell value: {value}"
            ))),
            Data::Error(value) => Err(ExcelError::Parse(format!(
                "cell contains Excel error: {value}"
            ))),
        }
    }
}
