# Examples

The workspace root is a virtual Cargo workspace, so runnable examples live under
the `excelx` package.

Run them from the workspace root:

```sh
cargo run -p excelx --example manual_read_write
cargo run -p excelx --example derive_read_write
cargo run -p excelx --example multi_sheet
```

Each example writes an `.xlsx` file to `target/examples/`.
