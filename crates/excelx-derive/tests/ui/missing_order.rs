use excelx_derive::ExcelRow;

#[derive(ExcelRow)]
struct MissingOrder {
    #[excel(header = "ID")]
    id: i64,
}

fn main() {}
