/// Single-sheet write options for Phase 1.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SheetOptions {
    pub name: String,
}

impl Default for SheetOptions {
    fn default() -> Self {
        Self {
            name: "Sheet1".to_owned(),
        }
    }
}

/// Homogeneous sheet data for multi-sheet workbook generation.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SheetData<T> {
    pub name: String,
    pub rows: Vec<T>,
}

impl<T> SheetData<T> {
    pub fn new(name: impl Into<String>, rows: Vec<T>) -> Self {
        Self {
            name: name.into(),
            rows,
        }
    }
}
