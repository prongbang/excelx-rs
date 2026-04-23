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
