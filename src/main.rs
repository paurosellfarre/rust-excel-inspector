#![windows_subsystem = "windows"]
use iced::widget::{button, column, container, text};
use iced::{Element, Sandbox, Settings};


use native_dialog::FileDialog;

use std::path::PathBuf;

mod xlsx_reader;
mod xlsx_writer;
  
#[derive(Default)]
pub struct MainFrame {
  paths: Vec<PathBuf>,
  opened_file: Option<PathBuf>,
  error_message: Option<String>,
  monthly_employee_count: Vec<i32>,
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
          self.monthly_employee_count = vec![0; 12];
          self.error_message = None;

          for path in &self.paths {
            //Check if opened is a file or a directory
            if path.is_dir() {
              self.error_message = Some(format!("Error: You have not selected an Excel file"));
              println!("Error: {} is a directory", path.to_string_lossy());
              continue;
            }
  
            self.opened_file = Some(path.to_path_buf());
  
            match xlsx_reader::read_xlsx_file(&path.to_path_buf(), &mut self.monthly_employee_count) {
              
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

          //If there was no files processed, do not create the resume file
          if self.monthly_employee_count.iter().sum::<i32>() > 0 {
            if let Err(err) = xlsx_writer::write_resume_xsls_file(&self.monthly_employee_count) {
              let new_error = format!("\nError writing resume xlsx file: {}", err);
              self.error_message = match &self.error_message {
                Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                _ => Some(new_error),
              };
            } else {
              let new_error = format!("\nFile resumen.xlsx processed successfully");
              self.error_message = match &self.error_message {
                Some(s) if !s.is_empty() => Some(format!("{}. {}", s, new_error)),
                _ => Some(new_error),
              };
            }
          }
        }
      }


    
  }

  fn view(&self) -> Element<Self::Message> {
    let open_button = button::Button::new( text::Text::new("Open File"))
      .on_press(Message::OpenFile);

    let content = column![
      text::Text::new("Excel Inspector. Select your files"),
      open_button,
      match &self.error_message {
        Some(s) => text::Text::new(s),
        None => text::Text::new(""),
      }
    ]
    .spacing(20);

    let container = container::Container::new(content)
      .width(iced::Length::Fill)
      .height(iced::Length::Fill)
      .center_x()
      .center_y();

    container.into()
    
  }
  

}

fn main() {

  let error = MainFrame::run(Settings{
    window: iced::window::Settings {
        size: iced::Size::new(300.0, 300.0),
        ..Default::default()
    },
    ..Default::default()
  });

  if error.is_err() {
    eprintln!("{}", error.unwrap_err());
  }
}