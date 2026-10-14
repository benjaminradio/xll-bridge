//! Embeds files and folders into a workbook (as a custom XML part) and inspects/extracts them.
//!
//!   xllb-pack embed   <in.xlsx> <out.xlsx> <file-or-folder>...   (folders keep their name)
//!   xllb-pack list    <book.xlsx>
//!   xllb-pack extract <book.xlsx> <dest-folder>
//!
//! Files given on the command line land in the workbook's root, so a `.lua` file runs when the
//! workbook is enabled; a folder `lib/` becomes `lib/` and is reachable through `require`.

use std::path::Path;
use std::process::ExitCode;

use xll_bridge::package::{embed_bundle, read_bundle};
use xll_bridge::vfs::{Bundle, Kind};

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match run(&args) {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("error: {e}");
            ExitCode::FAILURE
        }
    }
}

fn run(args: &[String]) -> Result<(), Box<dyn std::error::Error>> {
    match args.iter().map(String::as_str).collect::<Vec<_>>().as_slice() {
        ["embed", input, output, items @ ..] if !items.is_empty() => {
            // Start from what the workbook already carries so repeated embeds accumulate.
            let mut bundle = read_bundle(Path::new(input))?.unwrap_or_default();
            for item in items {
                let report = bundle.import_path(Path::new(item), "", true)?;
                for s in &report.skipped {
                    eprintln!("skipped: {s}");
                }
            }
            embed_bundle(Path::new(input), Path::new(output), &bundle)?;
            println!(
                "wrote {output} ({} file(s), {} bytes, digest {})",
                bundle.file_count(),
                bundle.total_bytes(),
                &bundle.digest()[..12]
            );
            Ok(())
        }
        ["list", book] => {
            match read_bundle(Path::new(book))? {
                None => println!("no embedded files"),
                Some(b) => {
                    println!("digest {}", b.digest());
                    for e in b.entries() {
                        let depth = e.path.matches('/').count();
                        let name = e.path.rsplit('/').next().unwrap_or(&e.path);
                        match e.kind {
                            Kind::Dir => println!("{}{name}/", "  ".repeat(depth)),
                            Kind::File => println!("{}{name} ({} bytes)", "  ".repeat(depth), e.size),
                        }
                    }
                }
            }
            Ok(())
        }
        ["extract", book, dest] => {
            let b = read_bundle(Path::new(book))?.unwrap_or_else(Bundle::new);
            std::fs::create_dir_all(dest)?;
            let top: Vec<String> = b
                .entries()
                .into_iter()
                .filter(|e| !e.path.contains('/'))
                .map(|e| e.path)
                .collect();
            for t in &top {
                b.export_to(t, Path::new(dest))?;
            }
            println!("extracted {} top-level item(s) to {dest}", top.len());
            Ok(())
        }
        _ => Err("usage: xllb-pack embed <in.xlsx> <out.xlsx> <file-or-folder>... | list <book.xlsx> | extract <book.xlsx> <dest>".into()),
    }
}
