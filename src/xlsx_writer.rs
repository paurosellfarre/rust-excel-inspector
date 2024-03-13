use xlsxwriter::{Workbook, Format};

pub fn write_xlsx_file(data: &Vec<String>) -> Result<(), xlsxwriter::XlsxError> {

    let filename = data[0].clone() + "-resumen.xlsx";
   
    let workbook = Workbook::new(&filename)?;
    let mut sheet1 = workbook.add_worksheet(None)?;
    let mut format = Format::new();
    format.set_text_wrap();
    format.set_bold();

    sheet1.set_column(0, 0, 30.0, None)?; // Ajusta el ancho de la primera columna
    sheet1.set_column(1, 1, 30.0, None)?;
    sheet1.set_column(2, 2, 30.0, None)?;

    sheet1.write_string(0, 0, "Empresa", Some(&format))?;
    sheet1.write_string(0, 1, &data[0], None)?;

    sheet1.write_string(1, 0, "Dias previstos extraccion.", Some(&format))?;
    sheet1.write_string(1, 1, &data[1], None)?;

    sheet1.write_string(2, 0, "Numero de empleados", Some(&format))?;
    sheet1.write_string(2, 1, &data[2], None)?;

    workbook.close()?;

    Ok(())
}

pub fn write_resume_xsls_file(data: &Vec<i32>) -> Result<(), xlsxwriter::XlsxError> {

    let filename = "resumen.xlsx";
    let months = vec!["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
   
    let workbook = Workbook::new(&filename)?;
    let mut sheet1 = workbook.add_worksheet(None)?;
    let mut format = Format::new();
    format.set_text_wrap();

    sheet1.set_column(0, 0, 25.0, None)?; // Ajusta el ancho de la primera columna
    sheet1.set_column(1, 1, 25.0, None)?;

    sheet1.write_string(0, 0, "Mes", Some(Format::new().set_bold()))?;
    sheet1.write_string(0, 1, "Numero de empleados", Some(Format::new().set_bold()))?;

    for (index, value) in data.iter().enumerate() {
        sheet1.write_string(index as u32 + 1, 0, months[index], Some(&format))?;
        sheet1.write_number(index as u32 + 1, 1, *value as f64, Some(&format))?;
    }

    sheet1.write_string(13, 0, "Total", Some(Format::new().set_bold()))?;
    sheet1.write_number(13, 1, data.iter().sum::<i32>() as f64, Some(Format::new().set_bold()))?;

    workbook.close()?;

    Ok(())
}
