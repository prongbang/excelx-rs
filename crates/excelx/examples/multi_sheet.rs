use std::fs;
use std::path::PathBuf;

use excelx::{SheetData, to_xlsx_multi};

#[derive(Debug, PartialEq, excelx_derive::ExcelRow)]
struct InventoryItem {
    #[excel(header = "SKU", order = 1)]
    sku: String,
    #[excel(header = "Name", order = 2)]
    name: String,
    #[excel(header = "Quantity", order = 3, default = "0")]
    quantity: i64,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let in_stock = vec![
        InventoryItem {
            sku: "A-100".to_owned(),
            name: "Keyboard".to_owned(),
            quantity: 24,
        },
        InventoryItem {
            sku: "A-200".to_owned(),
            name: "Mouse".to_owned(),
            quantity: 48,
        },
    ];
    let backordered = vec![InventoryItem {
        sku: "B-300".to_owned(),
        name: "Monitor".to_owned(),
        quantity: 0,
    }];

    let bytes = to_xlsx_multi(&[
        SheetData::new("In Stock", in_stock),
        SheetData::new("Backordered", backordered),
    ])?;

    let path = example_path("multi_sheet.xlsx")?;
    fs::write(&path, bytes)?;
    println!("wrote {}", path.display());

    Ok(())
}

fn example_path(file_name: &str) -> Result<PathBuf, std::io::Error> {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("../../target/examples");
    fs::create_dir_all(&path)?;
    path.push(file_name);
    Ok(path)
}
