use super::ExcelError;
use crate::employees::{Employee, Shift};
use calamine::{DataType, Reader, Sheets};
use chrono::{Date, Duration, NaiveDate, Utc};
use std::collections::HashMap;

pub fn parse_worksheet<'a>(
    workbook: &'a mut Sheets,
    sheet: &'a str,
    date_range: DateColumnRange,
) -> Result<Vec<Employee>, ExcelError> {
    let mut map: HashMap<usize, Employee> = HashMap::new();

    // Read whole worksheet data and provide some statistics
    if let Some(Ok(range)) = workbook.worksheet_range(sheet) {
        for (row, col, cell) in range.cells() {
            match cell {
                DataType::String(txt) => {
                    if col == 0 && !map.contains_key(&row) {
                        map.insert(row, Employee::new(txt.clone()));
                    }

                    if let Some(e) = map.get_mut(&row) {
                        match col {
                            1 => e.overtime_schedule = txt.clone(),
                            2 => e.dist_code = txt.clone(),
                            3 => e.exp_account = txt.clone(),
                            _ => {}
                        }
                    }
                }
                DataType::Int(i) => {
                    if let Some(e) = map.get_mut(&row) {
                        match col {
                            1 => e.overtime_schedule = i.to_string(),
                            2 => e.dist_code = i.to_string(),
                            3 => e.exp_account = i.to_string(),
                            _ => {}
                        }
                    }
                }
                DataType::Float(f) => {
                    if let Some(e) = map.get_mut(&row) {
                        if date_range.in_range(col) && f > &0.0 {
                            e.hours.push(Shift {
                                col,
                                duration: Duration::hours((24.0 * *f) as i64),
                                date: date_range.date_from_column(col).unwrap(),
                            })
                        }

                        match col {
                            1 => e.overtime_schedule = f.round().to_string(),
                            2 => e.dist_code = f.round().to_string(),
                            3 => e.exp_account = f.round().to_string(),
                            _ => {}
                        }
                    }
                }
                _ => {}
            }
        }

        Ok(map.values().cloned().collect())
    } else {
        Err(ExcelError::WorksheetNotFound(sheet.to_string()))
    }
}

pub fn parse_date_range<'a>(
    workbook: &'a mut Sheets,
    sheet: &'a str,
) -> Result<DateColumnRange, ExcelError> {
    let result = workbook
        .worksheet_range(sheet)
        .ok_or_else(|| ExcelError::WorksheetNotFound(sheet.to_string()))?;
    let range = match result {
        Ok(r) => r,
        Err(e) => match e {
            calamine::Error::Io(err) => {
                return Err(ExcelError::Io(err));
            }
            calamine::Error::Xlsx(err) => {
                return Err(ExcelError::Xlsx(err));
            }
            calamine::Error::Msg(msg) => {
                return Err(ExcelError::Msg(msg.to_string()));
            }
            _ => {
                return Err(ExcelError::Unexpected);
            }
        },
    };

    // Read whole worksheet data and provide some statistics
    let mut date_range = DateColumnRange::new();
    for (row, col, cell) in range.cells() {
        let date = Date::<Utc>::from_utc(NaiveDate::from_ymd(1900, 1, 1), Utc);
        if let DataType::Float(f) = cell {
            // The date using the number of days since Jan 1, 1900 as provided.
            // You might be wondering why excel would do such a thing?
            // Too bad!
            let parsed_date = date + Duration::days((f.round() - 2.0) as i64);
            // Set date row to first row encountered with dates
            if date_range.row == None {
                date_range.row = Some(row);
                date_range.head = col;
                date_range.tail = col;
                date_range.start = Some(parsed_date);
            } else if date_range.row == Some(row) {
                date_range.tail = col;
                date_range.end = Some(parsed_date);
            }
        }
    }

    Ok(date_range)
}

#[derive(Debug, Clone)]
pub struct DateColumnRange {
    row: Option<usize>,
    index: usize,

    pub head: usize,
    pub tail: usize,
    pub start: Option<Date<Utc>>,
    pub end: Option<Date<Utc>>,
}

impl DateColumnRange {
    pub fn new() -> Self {
        Self {
            row: None,
            head: 0,
            tail: 0,
            start: None,
            end: None,
            index: 0,
        }
    }

    pub fn range(&self) -> Option<(Date<Utc>, Date<Utc>)> {
        if self.start.is_none() || self.end.is_none() {
            return None;
        }

        Some((self.start.unwrap(), self.end.unwrap()))
    }

    pub fn in_range(&self, col: usize) -> bool {
        if self.range().is_none() {
            return false;
        }

        if col < self.head || col > self.tail {
            return false;
        }

        true
    }

    pub fn date_from_column(&self, col: usize) -> Option<Date<Utc>> {
        if !self.in_range(col) {
            return None;
        }

        let n = col - self.head;

        Some(self.start.unwrap() + Duration::days(n as i64))
    }

    pub fn len(&self) -> usize {
        self.tail - self.head
    }
}

impl Iterator for DateColumnRange {
    // we will be counting with usize
    type Item = Date<Utc>;

    // next() is the only required method
    fn next(&mut self) -> Option<Self::Item> {
        let date = self.date_from_column(self.index + self.head);
        if date.is_some() {
            // Increment our count. This is why we started at zero.
            self.index += 1;
        }

        date
    }
}
