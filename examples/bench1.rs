use calamine::{open_workbook, Reader, Xlsx};

fn main() {
    // Open workbook
    let file = "/home/martin/Documents/NYC_311_SR_2010-2020-sample-1M-2007-365.xlsx";
    let mut excel: Xlsx<_> = open_workbook(file) //"/home/martinn/Documents/NYC_311_SR_2010-2020-sample-1M.xlsx")
        .expect("failed to find file");

    // Get worksheet
    let sheet = excel
        .worksheet_range("NYC_311_SR_2010-2020-sample-1M")
        .unwrap()
        .unwrap();
    println!("width: {}, height: {}", sheet.width(), sheet.height());
    // iterate over rows
    for _row in sheet.rows() {}
}
