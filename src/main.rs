mod excel;
mod utils;

use excel::*;
use utils::get_date_and_time;

use flexi_logger::{detailed_format, Duplicate, FileSpec, Logger};
use log::{error, info, warn};

use std::path::Path;

const INPUT_XLXS: &str = "test.xlsx";
const OUTPUT_XLXS: &str = "output.xlsx";
const SHEETNAME: &str = "Sheet1";

fn main() {
    // initialize the logger
    let _logger = Logger::try_with_str("info") // log warn and error
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

    let mut book = read_excel(Path::new(OUTPUT_XLXS)).unwrap();
    read_value(&book, SHEETNAME, "A1").unwrap();
    let max: (u32, u32) = get_max_col_and_row(&book, SHEETNAME).unwrap();
    append_new_row(&mut book, SHEETNAME).unwrap();
    change_value(
        &mut book,
        SHEETNAME,
        max.0 - (max.0 - 1),
        max.1 + 1,
        "NEW_VALUE_HERE".to_string(),
    )
    .unwrap();
    match write_excel(&book, Path::new(OUTPUT_XLXS)) {
        Ok(_) => info!("Success"),
        Err(err) => error!("{err:?}"),
    }
}
