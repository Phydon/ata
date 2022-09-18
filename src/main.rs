extern crate umya_spreadsheet;
use umya_spreadsheet::{reader, writer};
use flexi_logger::{detailed_format, Duplicate, FileSpec, Logger};
use log::{info, warn, error};
use chrono::Local;

use std:: {
    path::Path,
    error::Error,
};

fn main() {
    // initialize the logger
    let _logger = Logger::try_with_str("info") // log info, warn and error
        .unwrap()
        .format_for_files(detailed_format) // use timestamp for every log
        .log_to_file(FileSpec::default().suppress_timestamp()) // no timestamps in the filename
        .append() // use only one logfile
        .duplicate_to_stderr(Duplicate::Info) // print infos, warnings and errors also to the console
        .start()
        .unwrap();

    let (date, time) = get_date_and_time();
    warn!("{:?}", date);
    error!("{:?}", time);

    match excel() {
        Ok(()) => info!("Success"),
        Err(err) => error!("{err}"),
    }
}

pub fn get_date_and_time() -> (String, String) {
    // panics when adding '%.3f' or something similar
    // -> chrono error
    let date = Local::now().format("%d.%m.%Y").to_string();
    let time = Local::now().format("%R").to_string();

    (date, time)
}

pub fn excel() -> Result<(), Box<dyn Error>> {
    let path = Path::new("./test.xlsx");
    let path2 = Path::new("./output.xlsx");
    let sheetname = "Sheet1";

    // READER
    let mut book = reader::xlsx::read(path).expect("Unable to read xlsx file: {path}");

    // read value in cell A1
    let a_one_value = book.get_sheet_by_name(sheetname)?.get_value("A1");
    println!("A1 = {}", a_one_value);

    // Change value
    book.get_sheet_by_name_mut(sheetname)?.get_cell_by_column_and_row_mut(&1, &2).set_value("WASDWASD");
    let a_two_value = book.get_sheet_by_name(sheetname)?.get_value("A2");
    println!("A2 = {}", a_two_value);

    // Max col and max row
    let max_col_and_row: (u32, u32) = book
        .get_sheet_by_name(sheetname)?
        .get_highest_column_and_row();

    let max_col = max_col_and_row.0;
    let max_row = max_col_and_row.1;
    info!("{:?}, {:?}", max_col, max_row);

    // insert rows
    book.insert_new_row(sheetname, &max_col, &max_row);

    // WRITER
    let _ = writer::xlsx::write(&book, path2);

    Ok(())
}
