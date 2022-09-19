extern crate umya_spreadsheet;
use umya_spreadsheet::{reader, writer::{self, xlsx::XlsxError}, Spreadsheet};

use log::{info, warn, error};

use std:: {
    path::Path,
    error::Error,
};

pub fn read_excel(path: &Path) -> Result<Spreadsheet, Box<dyn Error>> {
    let book = reader::xlsx::read(path).expect("Unable to read xlsx");

    Ok(book)
}

pub fn read_value(book: &Spreadsheet, sheetname: &str, val: &str) -> Result<String, Box<dyn Error>> {
    let result = book.get_sheet_by_name(sheetname)?.get_value(val);

    Ok(result)
}

pub fn get_max_col_and_row(book: &Spreadsheet, sheetname: &str) -> Result<(u32, u32), Box<dyn Error>> {
    let max = book.get_sheet_by_name(sheetname)?.get_highest_column_and_row();

    Ok(max)
}

pub fn append_new_row(book: &mut Spreadsheet, sheetname: &str) -> Result<(), Box<dyn Error>> {
    let max_col_and_row: (u32, u32) = get_max_col_and_row(book, sheetname)?;

    let max_col = max_col_and_row.0;
    let max_row = max_col_and_row.1;
    info!("{:?}, {:?}", max_col, max_row);

    book.insert_new_row(sheetname, &max_col, &max_row);

    Ok(())
}

pub fn change_value<'a>(book: &'a mut Spreadsheet, sheetname: &'a str, col: u32, row: u32, new_val: String) -> Result<&'a Spreadsheet, Box<dyn Error>> {
    book.get_sheet_by_name_mut(sheetname)?.get_cell_by_column_and_row_mut(&col, &row).set_value(new_val);

    Ok(book)
}

pub fn write_excel(book: &Spreadsheet, path: &Path) -> Result<(), XlsxError> {
    let _ = writer::xlsx::write(&book, path)?;

    Ok(())
}
