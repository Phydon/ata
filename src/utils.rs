use chrono::Local;

pub fn get_date_and_time() -> (String, String) {
    // panics when adding '%.3f' or something similar
    // -> chrono error
    let date = Local::now().format("%d.%m.%Y").to_string();
    let time = Local::now().format("%R").to_string();

    (date, time)
}
