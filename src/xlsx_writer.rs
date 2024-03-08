use xlsxwriter::Workbook;

pub fn write_xlsx_file(data: &Vec<String>) -> Result<(), xlsxwriter::XlsxError> {

    let filename = data[0].clone() + "-resumen.xlsx";
   
    let workbook = Workbook::new(&filename)?;
    let mut sheet1 = workbook.add_worksheet(None)?;

    sheet1.write_string(0, 0, "Empresa", None)?;
    sheet1.write_string(0, 1, &data[0], None)?;

    sheet1.write_string(1, 0, "Dias previstos extraccion.", None)?;
    sheet1.write_string(1, 1, &data[1], None)?;

    sheet1.write_string(2, 0, "Numero de empleados", None)?;
    sheet1.write_string(2, 1, &data[2], None)?;

    workbook.close()?;

    Ok(())
}