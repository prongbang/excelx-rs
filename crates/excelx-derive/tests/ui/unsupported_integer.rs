use excelx_derive::ExcelRow;

#[derive(ExcelRow)]
struct UnsupportedInteger {
    #[excel(header = "ID", order = 1)]
    id: u64,
}

fn main() {}
