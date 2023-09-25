use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::borrow::BorrowMut;
use std::cell::{Cell, RefCell};
use std::cmp::max;
use std::env;
use std::fmt::Formatter;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file =
        "/home/martin/repos/studio/viz/services/file-manager/src/test/resources/kortisMessy.xlsx"
            .to_string();
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
        .expect("Expecting a sheet number as second argument");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let dest = sce.with_extension("csv");
    let mut dest = BufWriter::new(File::create(dest).unwrap());
    let mut xl = open_workbook_auto(&sce).unwrap();

    let mut table: Vec<RefCell<Vec<DataType>>> = Vec::new();
    let mut cells: Vec<((u32, u32), DataType)> = Vec::new();

    table.push(RefCell::new(Vec::new()));
    let range = xl.worksheet2(sheet, &mut |pos, data_type| {
        let (row, col) = (pos.0 as usize, pos.1 as usize);

        cells.push((pos, data_type.clone()));
        if (col >= table.len()) {
            for _ in 0..=col - table.len() {
                table.push(RefCell::new(Vec::new()));
            }
        }
        let mut c1 = table.get(col);
        let mut c2 = c1.expect(format!("No value at {col}").as_str());
        // and_then(|cell: &mut Cell<Vec<DataType>>| {
        let mut c2 = c2.borrow_mut();
        for _ in c2.len()..row {
            println!("Pushing value: Empty to pos: {}x{}", row - 1, col);
            c2.push(DataType::Empty);
        }
        println!("Pushing value: {} to pos: {:?}", data_type, pos);

        c2.borrow_mut().push(data_type);
        // Some(col.push(data_type))
        // });
    });
    let cols = table.len();
    let rows = table.get(0).unwrap().borrow().len();
    println!("columns: {}, rows: {}", cols, rows);
    #[derive(Debug, Clone, Default)]
    struct Types {
        bools: u32,
        emptys: u32,
        errors: u32,
        ints: u32,
        floats: u32,
        strings: u32,
        datetimes: u32,
        iso_datetimes: u32,
        durs: u32,
        iso_durs: u32,
    }
    impl Types {
        fn inc(&mut self, data_type: &DataType) {
            match data_type {
                DataType::Empty => self.emptys += 1,
                DataType::Bool(_) => self.bools += 1,
                DataType::Error(_) => self.errors += 1,
                DataType::DateTime(_) => self.datetimes += 1,
                DataType::DateTimeIso(_) => self.iso_datetimes += 1,
                DataType::String(_) => self.strings += 1,
                DataType::Int(_) => self.ints += 1,
                DataType::Float(_) => self.floats += 1,
                DataType::Duration(_) => self.durs += 1,
                DataType::DurationIso(_) => self.iso_durs += 1,
            }
        }
    }
    impl std::fmt::Display for Types {
        fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
            let mut stri = String::new();
            if (self.emptys > 0) {
                stri.push_str(format!("Empty: {}\n", self.emptys).as_str());
            }
            if (self.bools > 0) {
                stri.push_str(format!("Booleans: {}\n", self.bools).as_str());
            }
            if (self.errors > 0) {
                stri.push_str(format!("Errors: {}\n", self.errors).as_str());
            }
            if (self.datetimes > 0) {
                stri.push_str(format!("Datetimes: {}\n", self.datetimes).as_str());
            }
            if (self.iso_datetimes > 0) {
                stri.push_str(format!("ISO Datetimes {}\n", self.iso_datetimes).as_str());
            }
            if (self.strings > 0) {
                stri.push_str(format!("Strings: {}\n", self.strings).as_str());
            }
            if (self.ints > 0) {
                stri.push_str(format!("Ints: {}\n", self.ints).as_str());
            }
            if (self.floats > 0) {
                stri.push_str(format!("Floats: {}\n", self.floats).as_str());
            }
            if (self.durs > 0) {
                stri.push_str(format!("Durations: {}\n", self.durs).as_str());
            }
            if (self.iso_durs > 0) {
                stri.push_str(format!("ISO Dirations: {}\n", self.iso_durs).as_str());
            }
            // stri
            f.write_str(stri.as_str())
        }
    }
    let mut total_types = Types::default();
    for cols in table.iter() {
        let mut types = Types::default();

        for cells in cols.borrow().iter() {
            types.inc(cells);
            total_types.inc(cells);
        }
        println!("Column {}: {}", cols.borrow().get(0).or(Some(&DataType::String("Unknown".to_string()))).unwrap(), types);
    }
    for row in 0..rows {
        let mut stri = String::new();
        for col in 0..cols {
            stri.push_str(
                table
                    .get(col)
                    .unwrap()
                    .borrow()
                    .get(row)
                    .map_or(String::new(), |s| s.to_string())
                    .as_str(),
            );
            if (col < cols - 1) {
                stri.push_str("\t");
            }
        }
        println!("{}", stri);
    }
    //write_range(&mut dest, &range).unwrap();
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
