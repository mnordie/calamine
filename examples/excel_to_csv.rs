use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists
    let file = "/home/martin/Documents/NYC_311_SR_2010-2020-sample-1M-2007-365.xlsx".to_string();
    // let file = "/home/martin/Documents/HealthCare_OPLS_X_Y_singlewords.xlsx".to_string();
    // let file = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let file = env::args()
        .nth(1)
        .or(Some(file))
        .expect("Please provide an excel file to convert");
    let sheet = env::args()
        .nth(2)
        .or(Some("0".to_string()))
        .and_then(|s| Some(str::parse::<usize>(&*s)))
        .unwrap()
        .expect("Expecting a sheet name as second argument");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let dest = sce.with_extension("csv");
    let mut dest = BufWriter::new(File::create(dest).unwrap());
    let mut xl = open_workbook_auto(&sce).unwrap();
    let range = xl.worksheet_range_at(sheet).unwrap().unwrap();

    write_range(&mut dest, &range).unwrap();
}

fn write_range<W: Write>(dest: &mut W, range: &Range<DataType>) -> std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                DataType::Empty => Ok(()),
                DataType::String(ref s)
                | DataType::DateTimeIso(ref s)
                | DataType::DurationIso(ref s) => write!(dest, "{}", s),
                DataType::Float(ref f) | DataType::DateTime(ref f) | DataType::Duration(ref f) => {
                    write!(dest, "{}", f)
                }
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ";")?;
            }
        }
        write!(dest, "\r\n")?;
    }
    Ok(())
}
