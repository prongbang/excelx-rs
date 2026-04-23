use excelx_derive::ExcelRow;

#[derive(ExcelRow)]
struct NestedOption {
    #[excel(header = "Value", order = 1)]
    value: Option<Option<String>>,
}

fn main() {}
