use iced::widget::{button, column, text,  Column};
use iced::{Sandbox, Element, Settings};


use native_dialog::FileDialog;

use std::path::PathBuf;

mod xlsx_reader;
mod xlsx_writer;
  
#[derive(Default)]
pub struct MainFrame {
  paths: Vec<PathBuf>,
  opened_file: Option<PathBuf>,
  error_message: Option<String>,
}

#[derive(Debug, Clone)]
pub enum Message {
    OpenFile,
    FileOpened,
}


impl Sandbox for MainFrame {
  type Message = Message;

  fn new() -> Self {
    Self::default()
  }

  fn title(&self) -> String {
      String::from("MainFrame")
  }

  fn update(&mut self, _message: Self::Message) {

    match _message {
      Message::OpenFile => {

          // Show only files with the extension "xls/xlsx".
          let paths = FileDialog::new()
          .set_location("~/Desktop")
          .add_filter("Excel File", &["xls", "xlsx"])
          .show_open_multiple_file()
          .unwrap();

          self.paths = paths;
          self.update(Message::FileOpened);

        }

        Message::FileOpened => {

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
        }
      }


    
  }

  fn view(&self) -> Element<Self::Message> {
    
    let open_button = button::Button::new( text::Text::new("Open File"))
      .on_press(Message::OpenFile);

    let content = Column::new()
      .push(open_button);

    content.into()
    
  }
  

}

extern "system" {
  fn SetStdHandle(nStdHandle: u32, hHandle: *mut ()) -> i32;
}

fn main() {
  use std::os::windows::io::AsRawHandle;
  //Desktop route
  let error_log = r"C:\Users\Public\error_log.txt";
  let f = std::fs::File::create(error_log).unwrap();
  unsafe {
      SetStdHandle((-12_i32) as u32, f.as_raw_handle().cast())
  };

  let error = MainFrame::run(Settings::default());

  if error.is_err() {
    eprintln!("{}", error.unwrap_err());
  }
}