use std::convert::TryInto;

use crate::employees::{sum_of_hours, Employee, Shift};
use crate::excel::timecards::DateColumnRange;
use thiserror::Error;
use xlsxwriter::{Workbook, XlsxError};

#[derive(Error, Debug)]
pub enum ExcelWriteError {
    #[error("unexpected error occurred")]
    Unexpected,
    #[error("xlsx error")]
    Xlsx(XlsxError),
}

const DATE_FORMAT: &str = "%Y-%m-%d";

const TIMECARD_HEADER_HEADERS: &[&str] = &[
    "EMPLOYEE",
    "PEREND",
    "TIMECARD",
    "TCARDDESC",
    "TIMESLATE",
    "REUSECARD",
    "ACTIVE",
    "SEPARATECK",
    "PROCESSED",
    "CREGHRS",
    "CSHIFTHRS",
    "CVACHRSP",
    "CVACHRSA",
    "CSICKHRSP",
    "CSICKHRSP",
    "CCOMPHRSP",
    "CCOMPHRSA",
    "CVACAMTP",
    "CVACAMTA",
    "CSICKAMTP",
    "CSICKAMTA",
    "CCOMPAMTP",
    "CCOMPAMTA",
    "CDISIHRSP",
    "CDISIHRSA",
    "CDISIAMTP",
    "CDISIAMTA",
    "LASTNAME",
    "FIRSTNAME",
    "MIDDLENAME",
    "GREGHRS",
    "GSHIFTHRS",
    "GVACHRSP",
    "GVACHRSA",
    "GSICKHRSP",
    "GSICKHRSA",
    "GCOMPHRSP",
    "GCOMPHRSA",
    "GVACAMTP",
    "GVACAMTA",
    "GSICKAMTP",
    "GSICKAMTA",
    "GCOMPAMTP",
    "GCOMPAMTA",
    "KEYACTION",
    "GDISIHRSP",
    "GDISIHRSA",
    "GDISIAMTP",
    "GDISIAMTA",
    "HIREDATE",
    "FIREDATE",
    "PARTTIME",
    "PAYFREQ",
    "OTSCHED",
    "COMPTIME",
    "SHIFTSCHED",
    "SHIFTNUM",
    "WORKPROV",
    "STATUS",
    "INACTDATE",
    "PROCESSCMD",
    "GOTHOURS",
    "OTCALCTYPE",
    "HRSPERDAY",
    "WORKCODE",
    "TOTALJOBS",
    "USERSEC",
    "WKLYFLSA",
    "VALUES",
    "OTOVERRIDE",
    "COTHOURS",
    "TCDLINES",
    "SWJOB",
    "SRCEAPPL",
];

const TIMECARD_DETAIL_HEADERS: &[&str] = &[
    "EMPLOYEE",
    "PEREND",
    "TIMECARD",
    "LINENUM",
    "CATEGORY",
    "EARNDED",
    "EARDEDTYPE",
    "EARDEDDATE",
    "STARTTIME",
    "STOPTIME",
    "GLSEG1",
    "GLSEG2",
    "GLSEG3",
    "HOURS",
    "CALCMETH",
    "LIMITBASE",
    "CNTBASE",
    "RATE",
    "PAYORACCR",
    "EXPACCT",
    "LIABACCT",
    "OTACCT",
    "SHIFTACCT",
    "ASSETACCT",
    "OTSCHED",
    "SHIFTSCHED",
    "SHIFTNUM",
    "WCC",
    "TAXWEEKS",
    "TAXANNLIZ",
    "WEEKLYNTRY",
    "ENTRYTYPE",
    "POOLEDTIPS",
    "DESC",
    "GLSEGID1",
    "GLSEGDESC1",
    "GLSEGID2",
    "GLSEGDESC2",
    "GLSEGID3",
    "GLSEGDESC3",
    "KEYACTION",
    "WORKPROV",
    "PROCESSCMD",
    "NKEMPLOYEE",
    "NKPEREND",
    "NKTIMECARD",
    "NKLINENUM",
    "DAYS",
    "WCCGROUP",
    "VALUES",
    "OTHOURS",
    "OTRATE",
    "SWFLSA",
    "DISTCODE",
    "REXPACCT",
    "RLIABACCT",
    "SWALLOCJOB",
    "JOBS",
    "WORKCODE",
    "JOBHOURS",
    "JOBBASE",
    "RCALCMETH",
    "RLIMITBASE",
    "RRATEOVER",
    "RRATE",
    "DEFRRATE",
];

pub fn generate(
    workbook: Workbook,
    payperiod: &str,
    employees: Vec<Employee>,
    date_range: DateColumnRange,
) -> Result<(), ExcelWriteError> {
    let mut sheet_header = workbook
        .add_worksheet(Some("Timecard_Header"))
        .map_err(ExcelWriteError::Xlsx)?;
    for (i, heading) in TIMECARD_HEADER_HEADERS.iter().enumerate() {
        sheet_header
            .write_string(0, i as u16, heading, None)
            .map_err(ExcelWriteError::Xlsx)?;
    }

    let mut sheet_detail = workbook
        .add_worksheet(Some("Timecard_Detail"))
        .map_err(ExcelWriteError::Xlsx)?;
    for (i, heading) in TIMECARD_DETAIL_HEADERS.iter().enumerate() {
        sheet_detail
            .write_string(0, i as u16, heading, None)
            .map_err(ExcelWriteError::Xlsx)?;
    }

    let (_, end) = date_range.range().ok_or(ExcelWriteError::Unexpected)?;
    let end_formatted = end.format(DATE_FORMAT).to_string(); // 2021-05-08 12:00:00 AM

    let employees: Vec<Employee> = employees
        .into_iter()
        .filter(|e| sum_of_hours(e.hours.clone()) > 0.0)
        .collect();
    for (i, employee) in employees.iter().enumerate() {
        let row = i as u32 + 1;

        // Timecard_Header
        sheet_header
            .write_string(row, 0, &employee.id, None)
            .map_err(ExcelWriteError::Xlsx)?;
        sheet_header
            .write_string(row, 1, &end_formatted, None)
            .map_err(ExcelWriteError::Xlsx)?;
        sheet_header
            .write_string(row, 2, payperiod, None)
            .map_err(ExcelWriteError::Xlsx)?;
    }
    sheet_header
        .merge_range(
            0,
            0,
            employees.len().try_into().unwrap(),
            (TIMECARD_HEADER_HEADERS.len() - 1).try_into().unwrap(),
            "Timecard_Header",
            None,
        )
        .map_err(ExcelWriteError::Xlsx)?;

    let mut row = 1;
    for employee in employees {
        // Timecard_Detail
        let shifts: Vec<Shift> = employee
            .hours
            .clone()
            .into_iter()
            .filter(|shift| shift.sum_of_shift() > 0.0)
            .collect();
        for i in 0..shifts.len() {
            let shift = employee.hours[i];

            // A
            sheet_detail
                .write_string(row, 0, &employee.id, None)
                .map_err(ExcelWriteError::Xlsx)?;
            // B
            sheet_detail
                .write_string(row, 1, &end_formatted, None)
                .map_err(ExcelWriteError::Xlsx)?;
            // C
            sheet_detail
                .write_string(row, 2, payperiod, None)
                .map_err(ExcelWriteError::Xlsx)?;
            // D
            sheet_detail
                .write_string(row, 3, &format!("{}", (i + 1) * 1000), None)
                .map_err(ExcelWriteError::Xlsx)?;
            // E
            sheet_detail
                .write_string(row, 4, "2", None)
                .map_err(ExcelWriteError::Xlsx)?;
            // E
            sheet_detail
                .write_string(row, 5, "HRLY", None)
                .map_err(ExcelWriteError::Xlsx)?;
            // F
            sheet_detail
                .write_string(row, 5, "HRLY", None)
                .map_err(ExcelWriteError::Xlsx)?;
            // H
            sheet_detail
                .write_string(row, 7, &shift.date.format(DATE_FORMAT).to_string(), None)
                .map_err(ExcelWriteError::Xlsx)?;
            // N
            sheet_detail
                .write_string(row, 13, &format!("{}", shift.sum_of_shift()), None)
                .map_err(ExcelWriteError::Xlsx)?;
            // T
            sheet_detail
                .write_string(row, 19, &employee.exp_account, None)
                .map_err(ExcelWriteError::Xlsx)?;
            // V
            sheet_detail
                .write_string(row, 21, &employee.exp_account, None)
                .map_err(ExcelWriteError::Xlsx)?;
            // Y
            sheet_detail
                .write_string(row, 24, &employee.overtime_schedule, None)
                .map_err(ExcelWriteError::Xlsx)?;
            // AV
            sheet_detail
                .write_string(row, 47, "1", None)
                .map_err(ExcelWriteError::Xlsx)?;
            // BB
            sheet_detail
                .write_string(row, 53, &employee.dist_code, None)
                .map_err(ExcelWriteError::Xlsx)?;

            row += 1;
        }
    }

    sheet_detail
        .merge_range(
            0,
            0,
            row,
            (TIMECARD_DETAIL_HEADERS.len() - 1).try_into().unwrap(),
            "Timecard_Detail",
            None,
        )
        .map_err(ExcelWriteError::Xlsx)?;

    workbook.close().map_err(ExcelWriteError::Xlsx)?;

    Ok(())
}
