use excelx_derive::ExcelRow;

#[derive(ExcelRow)]
struct UnsupportedType {
    #[excel(header = "Tags", order = 1)]
    tags: Vec<String>,
}

fn main() {}
