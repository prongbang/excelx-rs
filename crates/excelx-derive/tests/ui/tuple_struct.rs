use excelx_derive::ExcelRow;

#[derive(ExcelRow)]
struct TupleStruct(#[excel(header = "ID", order = 1)] i64);

fn main() {}
