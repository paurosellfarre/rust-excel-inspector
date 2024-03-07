// xlsx_reader.rs

use calamine::{open_workbook, Reader, Xlsx};
use std::path::PathBuf;

pub fn read_xlsx_file(path: &PathBuf) -> Vec<String> {
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
    let mut results: Vec<String> = Vec::new();

    //TODO: Refactor into a function to handle additional future sheets in an easier way
    //From first sheet we want
    //Empresa, Dias previstos extraccion.
    let firt_sheet_expected = vec!["Empresa", "Dias previstos extraccion."];
    
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
        for row in range.rows() {
            println!("Row: {:?}", row);

            //Check if the first cell of the row is the expected one
            if let Some(cell) = row.get(0) {
                for expected in &firt_sheet_expected {
                    if cell.to_string() == *expected {
                        
                        //Get the next cell
                        if let Some(next_cell) = row.get(1) {
                            println!("Next cell: {:?}", next_cell);
                            results.push(next_cell.to_string());
                        }
                    }
                }
            }
            
        }
    }

    //From second sheet we want to count the number of DNIs (column 2)
    //so we know how many employees are in the company
    //DNI example: 45735359D

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
    }

    results
}