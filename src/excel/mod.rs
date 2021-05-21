pub mod sage;
pub mod timecards;

use thiserror::Error;
use unicode_segmentation::UnicodeSegmentation;

const ASCII_LETTERS: &'static str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

#[derive(Error, Debug)]
pub enum ExcelError {
    #[error("unexpected error occurred")]
    Unexpected,
    #[error("io error")]
    Io(std::io::Error),
    #[error("xlsx error")]
    Xlsx(calamine::XlsxError),
    #[error("{0}")]
    Msg(String),
    #[error("worksheet not found '{0}'")]
    WorksheetNotFound(String),
}

pub fn to_column_letter(col: i32) -> String {
    let mut string = String::from("");
    let mut n = col + 1;

    while n > 0 {
        let rem = (n % 26) as u8;

        if rem == 0 {
            string.push('Z');
            n = (n / 26) - 1;
        } else {
            string.push(((rem - 1) + 'A' as u8) as char);
            n = n / 26;
        }
    }

    string.graphemes(true).rev().collect::<Vec<&str>>().concat()
}

pub fn from_column_letter(col: String) -> i32 {
    let mut num: i32 = 0;
    for c in col.chars() {
        if ASCII_LETTERS.contains(c) {
            num = num * 26 + (c as u8 - 'A' as u8) as i32 + 1
        }
    }

    num - 1
}

#[test]
fn it_converts_to_columns() {
    let mut index = 0;
    for letter in ASCII_LETTERS.chars() {
        assert_eq!(to_column_letter(index), letter.to_string());
        index += 1;
    }

    assert_eq!(to_column_letter(25 + 1), "AA");
    assert_eq!(to_column_letter(26 * 2 + 24), "BY");
}

#[test]
fn it_converts_from_columns() {
    assert_eq!(from_column_letter("AA".to_string()), 25 + 1);
    assert_eq!(from_column_letter("BY".to_string()), 26 * 2 + 24);
}
