// xlsx_reader.rs

use calamine::{open_workbook, Reader, Xlsx};
use std::{fmt::Error, path::PathBuf};
use chrono::{NaiveDate, Duration, Datelike};

pub fn read_xlsx_file(path: &PathBuf, monthly_employee_count: &mut Vec<i32>) -> Result<Vec<String>, Error> {
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
    let mut results: Vec<String> = Vec::new();
    let mut current_file_month: u32 = 13;

    //TODO: Refactor into a function to handle additional future sheets in an easier way
    //From first sheet we want
    //Empresa, Dias previstos extraccion.
    let firt_sheet_expected = vec!["Empresa", "Dias previstos extraccion."];
    
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {

        for row in range.rows() {
            let cell = match row.get(0) {
                Some(cell) => cell,
                None => continue,
            };

            let expected_index = firt_sheet_expected.iter().enumerate()
                .find(|(_, expected)| cell.to_string() == **expected)
                .map(|(index, _)| index);

            let expected_index = match expected_index {
                Some(index) => index,
                None => continue,
            };

            let next_cell = match row.get(1) {
                Some(next_cell) => next_cell,
                None => continue,
            };

            if expected_index == 1 {
                let cell_string = next_cell.to_string();
                println!("Cell string: {}", cell_string);
                let date = NaiveDate::parse_from_str(&cell_string, "%d/%m/%Y")
                    .or_else(|_| {
                        next_cell.to_string().parse::<i32>()
                            .map(|days_from_excel_epoch| {
                                let excel_epoch = NaiveDate::from_ymd(1899, 12, 30);
                                excel_epoch + Duration::days(days_from_excel_epoch as i64)
                            })
                    })
                    .ok();

                if let Some(date) = date {
                    println!("Date: {:?}", date);
                    current_file_month = date.month() - 1;
                    results.push(date.format("%d/%m/%Y").to_string());
                }
            } else {
                results.push(next_cell.to_string());
            }
        }
        
    }

    if results.len() != firt_sheet_expected.len() {
        println!("Error: Not all expected data was found in the file");
        println!("Results: {:?}", results);
        return Err(Error);
    }

    //From second sheet we want to count the number of DNIs (column 2)
    //so we know how many employees are in the company
    //DNI example: 45735359D
    let second_sheet_expected = vec!["DNI"];

    if let Some(Ok(range)) = workbook.worksheet_range_at(1) {
        let mut dni_count = 0;
        for row in range.rows() {
            if let Some(cell) = row.get(1) {
                if cell.to_string().len() > 0 {
                    dni_count += 1;
                }
            }
        }
        dni_count -= 1; //Substract the header
        results.push(dni_count.to_string());
        monthly_employee_count[current_file_month as usize] += dni_count;
        println!("DNI count: {}", dni_count);
        println!("Current file month: {}", current_file_month);
    }

    if results.len() != firt_sheet_expected.len() + second_sheet_expected.len() {
        println!("Error: Not all expected data was found in the file");
        println!("Results: {:?}", results);
        return Err(Error);
    }
    println!("Monthly employee count: {:?}", monthly_employee_count);

    Ok(results)
}