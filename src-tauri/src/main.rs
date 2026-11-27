// no console window behind the app on Windows
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    study_pacer_lib::run()
}
