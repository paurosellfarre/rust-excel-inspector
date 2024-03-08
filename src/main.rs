#![windows_subsystem = "windows"]
use eframe::{
    egui::{CentralPanel, Context, RichText, ViewportBuilder, Image, include_image},
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
        ui.add(Image::new(include_image!("../assets/inspector-image.png")).max_width(250.0));
        ui.add_space(110.0);

        ui.label("To select more than one file, hold Ctrl button and click on the files. Finally press Open.");

        ui.add_space(20.0);

        if (ui.button(RichText::new("Search Excels").size(20.0))).clicked() {
        // Show only files with the extension "xls/xlsx".
        let filter = Box::new({
            let ext_xls = Some(OsStr::new("xls"));
            let ext_xlsx = Some(OsStr::new("xlsx"));
            move |path: &Path| -> bool { path.extension() == ext_xls 
                || path.extension() == ext_xlsx }
        });

        let mut dialog = FileDialog::open_file(self.opened_file.clone())
          .show_files_filter(filter)
          .resizable(false)
          .show_new_folder(false)
          .show_rename(false)
          .show_drives(false)
          .multi_select(true);

        dialog.open();

        self.open_file_dialog = Some(dialog);
        }

        if let Some(error_message) = &self.error_message {
          ui.label(error_message);
        }

        if let Some(dialog) = &mut self.open_file_dialog {
          if dialog.show(ctx).selected() {
            self.error_message = Some("".to_string());

            for file in dialog.selection() {

              //Check if opened is a file or a directory
              if file.is_dir() {
                self.error_message = Some(format!("Error: You have not selected an Excel file"));
                println!("Error: {} is a directory", file.to_string_lossy());
                continue;
              }

              self.opened_file = Some(file.to_path_buf());

              match xlsx_reader::read_xlsx_file(&file.to_path_buf()) {
                
                Ok(results) => {
                  if let Err(err) = xlsx_writer::write_xlsx_file(&results) {

                    let new_error = format!("\nError writing xlsx file: {}", err);
                    self.error_message = match &self.error_message {
                      Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                      _ => Some(new_error),
                    };
                  } else {
                    let filename = file.file_name().unwrap().to_str().unwrap();
                    let new_error = format!("\nFile {} processed successfully", filename);
                    self.error_message = match &self.error_message {
                      Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                      _ => Some(new_error),
                    };
                  }
                }

                Err(_) => {
                  let filename = file.file_name().unwrap().to_str().unwrap();
                  let new_error = format!("\nError reading xlsx file: {}", filename);
                  self.error_message = match &self.error_message {
                    Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                    _ => Some(new_error),
                  };
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
      .with_inner_size([400.0, 588.0])
      .with_resizable(false),
      ..Default::default()
    },

    Box::new(|_cc| Box::new(MainFrame::default())),
  );
}