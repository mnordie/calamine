use calamine::{open_workbook_auto, CellPos, DataType, Dimension, Reader, SheetCallbacks};
use std::borrow::BorrowMut;
use std::cell::RefCell;
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

struct CsvWriter {
    table: Vec<String>,
    rows: u32,
    cols: u32,
    dims: Dimension,
    dest: Box<dyn Write>,
    last_row: u32,
}
impl CsvWriter {
    fn format_cell(&self, cell: &DataType) -> String {
        match *cell {
            DataType::Empty => format!(""),
            DataType::String(ref s) => {
                if s.contains([',', '"']) {
                    let s = s.replace(&['\"'][..], "\"\"");
                    format!("\"{s}\"")
                } else {
                    format!("{s}")
                }
            }

            DataType::DateTimeIso(ref s) | DataType::DurationIso(ref s) => format!("{}", s),
            DataType::Float(ref f) | DataType::DateTime(ref f) | DataType::Duration(ref f) => {
                format!("{}", f)
            }

            DataType::Int(ref i) => format!("{}", i),
            DataType::Error(ref e) => format!("{:?}", e),
            DataType::Bool(ref b) => format!("{}", b),
        }
    }
}
impl SheetCallbacks for CsvWriter {
    fn dimension(&mut self, dim: Dimension) {
        self.rows = dim.end.row - dim.start.row;
        self.cols = dim.end.col - dim.start.col;
        // self.dims = *dim; //.clone();
        // self.dims.end = dim.end;
        self.dims.start.row = dim.start.row;
        self.dims.start.col = dim.start.col;
        self.dims.end.row = dim.end.row;
        self.dims.end.col = dim.end.col;
        // println!("->> dim: {:?}", dim); // Dimension: Dimensions { start: (0, 0), end: (36, 4) }
        println!("->> self.dims: {:?}", self.dims); // Dimension: Dimensions { start: (0, 0), end: (36, 4) }
    }

    fn cell(&mut self, pos: CellPos, cell: DataType) {
        // println!("->> Cell: {:?}@{:?}", cell, pos);
        if pos.row > self.last_row {
            self.last_row = pos.row;
            self.row_end(pos);
            self.table.clear(); // = vec![];
        }
        self.table.push(self.format_cell(&cell));
    }

    fn row_end(&mut self, _pos: CellPos) {
        // println!("->> Row end  at {:?}", pos);
        // println!("->> Dimension: {:?}", self.dims); // Dimension: Dimensions { start: (0, 0), end: (36, 4) }
        for (i, cell) in self.table.iter().enumerate() {
            let _ = write!(self.dest, "{}", cell.as_str());
            // println!("cols:{}", self.cols);
            if i < self.table.len() - 1 {
                // println!("apa");
                let _ = write!(self.dest, ",");
            }
        }
        let _ = write!(self.dest, "\n");
    }
}
fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = "/home/martin/Documents/NYC_311_SR_2010-2020-sample-1M-2007-365.xlsx"
        // "/home/martin/repos/studio/viz/services/file-manager/src/test/resources/kortisMessy.xlsx"
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
    let dest = BufWriter::new(File::create(dest).unwrap());
    let mut xl = open_workbook_auto(&sce).unwrap();
    let mut sheet_handler = CsvWriter {
        table: vec![],
        rows: 0,
        cols: 0,
        dims: Dimension::default(),
        dest: Box::new(dest),
        last_row: 0,
    };
    let _range = xl.worksheet2(sheet, &mut sheet_handler);
    //write_range(&mut dest, &range).unwrap();
    // sheet_handler.find_headers();
}
