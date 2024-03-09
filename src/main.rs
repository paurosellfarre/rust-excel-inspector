#![windows_subsystem = "windows"]
use eframe::{
  egui::{CentralPanel, Context, RichText, ViewportBuilder, Image, include_image},
  App, Frame,
};

use native_dialog::FileDialog;

use egui_extras;

use std::path::PathBuf;

mod xlsx_reader;
mod xlsx_writer;
  
#[derive(Default)]
pub struct MainFrame {
  paths: Vec<PathBuf>,
  opened_file: Option<PathBuf>,
  error_message: Option<String>,
}
  
impl App for MainFrame {
  fn update(&mut self, ctx: &Context, _frame: &mut Frame) {
    CentralPanel::default().show(ctx, |ui| {

      ui.vertical_centered(|ui| {

        egui_extras::install_image_loaders(ctx);
        ui.add(Image::new(include_image!("../assets/inspector-image.png")).max_width(250.0));
        ui.add_space(110.0);

        if (ui.button(RichText::new("Search Excels").size(20.0))).clicked() {

          // Show only files with the extension "xls/xlsx".
          self.paths = FileDialog::new()
          .set_location("~/Desktop")
          .add_filter("Excel File", &["xls", "xlsx"])
          .show_open_multiple_file()
          .unwrap();

        }

        if let Some(error_message) = &self.error_message {
          ui.label(error_message);
        }

        self.error_message = Some("".to_string());

        for path in &self.paths {

          //Check if opened is a file or a directory
          if path.is_dir() {
            self.error_message = Some(format!("Error: You have not selected an Excel file"));
            println!("Error: {} is a directory", path.to_string_lossy());
            continue;
          }

          self.opened_file = Some(path.to_path_buf());

          match xlsx_reader::read_xlsx_file(&path.to_path_buf()) {
            
            Ok(results) => {
              if let Err(err) = xlsx_writer::write_xlsx_file(&results) {

                let new_error = format!("\nError writing xlsx file: {}", err);
                self.error_message = match &self.error_message {
                  Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                  _ => Some(new_error),
                };
              } else {
                let filename = path.file_name().unwrap().to_str().unwrap();
                let new_error = format!("\nFile {} processed successfully", filename);
                self.error_message = match &self.error_message {
                  Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                  _ => Some(new_error),
                };
              }
            }

            Err(_) => {
              let filename = path.file_name().unwrap().to_str().unwrap();
              let new_error = format!("\nError reading xlsx file: {}", filename);
              self.error_message = match &self.error_message {
                Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                _ => Some(new_error),
              };
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