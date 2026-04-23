use std::fs;
use std::path::PathBuf;

use excelx::{from_xlsx, to_xlsx};

#[derive(Debug, PartialEq, excelx_derive::ExcelRow)]
struct Account {
    #[excel(header = "ID", order = 1)]
    id: i64,
    #[excel(header = "Email", order = 2)]
    email: String,
    #[excel(header = "Active", order = 3, default = "true")]
    active: bool,
    #[excel(header = "Plan", order = 4, default = "free")]
    plan: Option<String>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let accounts = vec![
        Account {
            id: 1,
            email: "ada@example.com".to_owned(),
            active: true,
            plan: Some("pro".to_owned()),
        },
        Account {
            id: 2,
            email: "grace@example.com".to_owned(),
            active: false,
            plan: None,
        },
    ];

    let bytes = to_xlsx(&accounts)?;
    let parsed = from_xlsx::<Account>(&bytes)?;
    assert_eq!(
        parsed,
        vec![
            Account {
                id: 1,
                email: "ada@example.com".to_owned(),
                active: true,
                plan: Some("pro".to_owned()),
            },
            Account {
                id: 2,
                email: "grace@example.com".to_owned(),
                active: false,
                plan: Some("free".to_owned()),
            },
        ]
    );

    let path = example_path("derive_read_write.xlsx")?;
    fs::write(&path, bytes)?;
    println!("wrote {}", path.display());

    Ok(())
}

fn example_path(file_name: &str) -> Result<PathBuf, std::io::Error> {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("../target/examples");
    fs::create_dir_all(&path)?;
    path.push(file_name);
    Ok(path)
}
