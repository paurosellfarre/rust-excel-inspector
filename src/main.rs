use eframe::{
    egui::{CentralPanel, Context, RichText, ViewportBuilder},
    App, Frame,
};

use egui_file::FileDialog;
use egui_extras;

use std::{
    ffi::OsStr,
    path::{Path, PathBuf},
};

mod xlsx_reader;
mod xlsx_writer;
  
#[derive(Default)]
pub struct MainFrame {
  opened_file: Option<PathBuf>,
  open_file_dialog: Option<FileDialog>,
  error_message: Option<String>,
}
  
impl App for MainFrame {
  fn update(&mut self, ctx: &Context, _frame: &mut Frame) {
    CentralPanel::default().show(ctx, |ui| {

      ui.vertical_centered(|ui| {

        egui_extras::install_image_loaders(ctx);
        ui.image("file://assets/inspector-image.png");
        ui.add_space(110.0);

        if (ui.button(RichText::new("Choose Excels").size(20.0))).clicked() {
        // Show only files with the extension "xls/xlsx".
        let filter = Box::new({
            let ext_xls = Some(OsStr::new("xls"));
            let ext_xlsx = Some(OsStr::new("xlsx"));
            move |path: &Path| -> bool { path.extension() == ext_xls 
                || path.extension() == ext_xlsx }
        });

        let mut dialog = FileDialog::open_file(self.opened_file.clone())
          .show_files_filter(filter)
          .multi_select(true);

        dialog.open();

        self.open_file_dialog = Some(dialog);
        }

        if let Some(error_message) = &self.error_message {
          ui.label(error_message);
        }

        if let Some(dialog) = &mut self.open_file_dialog {
          if dialog.show(ctx).selected() {

            for file in dialog.selection() {
              self.opened_file = Some(file.to_path_buf());
              println!("Opened file: {:?}", file.to_path_buf());

              match xlsx_reader::read_xlsx_file(&file.to_path_buf()) {

                Ok(results) => { 
                  if let Err(err) = xlsx_writer::write_xlsx_file(&results) {
                    self.error_message = Some(format!("Error writing xlsx file: {}", err));
                  }
                }

                Err(err) => {
                  self.error_message = Some(format!("Error reading xlsx file: {}", err));
                }

              }  
            }

          }
        }

      });

        
      
    });
  
  }
}

fn main() {
  let _ = eframe::run_native(
    "Excel Inspector Gadget",
    eframe::NativeOptions {
      viewport: ViewportBuilder::default()
      .with_inner_size([400.0, 600.0]),
      ..Default::default()
    },

    Box::new(|_cc| Box::new(MainFrame::default())),
  );
}