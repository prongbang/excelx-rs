use excelx_derive::ExcelRow;

#[derive(ExcelRow)]
struct MissingHeader {
    #[excel(order = 1)]
    id: i64,
}

fn main() {}
