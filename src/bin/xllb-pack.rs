//! Embeds Lua scripts into a workbook as a custom XML part.
//!
//!   xllb-pack embed <in.xlsx> <out.xlsx> <script.lua>...   (script name = file stem)
//!   xllb-pack list  <book.xlsx>

use std::path::{Path, PathBuf};
use std::process::ExitCode;

use xll_bridge::package::{Script, ScriptSet, embed_scripts, read_scripts};

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
        ["embed", input, output, scripts @ ..] if !scripts.is_empty() => {
            let mut set = ScriptSet::default();
            for p in scripts {
                let path = PathBuf::from(p);
                let name = path.file_stem().and_then(|s| s.to_str()).ok_or("bad script file name")?;
                set.scripts.push(Script { name: name.to_owned(), code: std::fs::read_to_string(&path)? });
            }
            embed_scripts(Path::new(input), Path::new(output), &set)?;
            println!("wrote {output} ({} script(s), digest {})", set.scripts.len(), &set.digest()[..12]);
            Ok(())
        }
        ["list", book] => {
            match read_scripts(Path::new(book))? {
                None => println!("no scripts"),
                Some(set) => {
                    println!("digest {}", set.digest());
                    for s in set.scripts {
                        println!("- {} ({} bytes)", s.name, s.code.len());
                    }
                }
            }
            Ok(())
        }
        _ => Err("usage: xllb-pack embed <in.xlsx> <out.xlsx> <script.lua>... | xllb-pack list <book.xlsx>".into()),
    }
}
