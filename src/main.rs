use eframe::{
    egui::{CentralPanel, Context},
    App, Frame,
};

use egui_file::FileDialog;

use std::{
    ffi::OsStr,
    path::{Path, PathBuf},
};

mod xlsx_reader;
  
  #[derive(Default)]
  pub struct Demo {
    opened_file: Option<PathBuf>,
    open_file_dialog: Option<FileDialog>,
  }
  
  impl App for Demo {
    fn update(&mut self, ctx: &Context, _frame: &mut Frame) {

        CentralPanel::default().show(ctx, |ui| {

            if (ui.button("Open")).clicked() {
            // Show only files with the extension "xls/xlsx".
            let filter = Box::new({
                let ext_xls = Some(OsStr::new("xls"));
                let ext_xlsx = Some(OsStr::new("xlsx"));
                move |path: &Path| -> bool { path.extension() == ext_xls 
                    || path.extension() == ext_xlsx }
            });

            let mut dialog = FileDialog::open_file(self.opened_file.clone()).show_files_filter(filter);
            dialog.open();

            self.open_file_dialog = Some(dialog);
            }
    
            if let Some(dialog) = &mut self.open_file_dialog {
                if dialog.show(ctx).selected() {
                    
                    if let Some(file) = dialog.path() {
                        self.opened_file = Some(file.to_path_buf());
                        println!("Opened file: {:?}", file.to_path_buf());

                        let results = xlsx_reader::read_xlsx_file(&file.to_path_buf());
                        println!("Results: {:?}", results);
                    }
    
                }

            }
            
        });
    }
  }
  
  fn main() {
    let _ = eframe::run_native(
      "File Dialog Demo",
      eframe::NativeOptions::default(),

      Box::new(|_cc| Box::new(Demo::default())),
    );
  }