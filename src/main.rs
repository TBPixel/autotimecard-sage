#[macro_use]
extern crate log;
extern crate chrono;

use anyhow::Context;
use env_logger::Env;
use excel::timecards::DateColumnRange;
use structopt::StructOpt;

use crate::excel::{from_column_letter, to_column_letter};

mod employees;
mod excel;

const DATE_FORMAT: &'static str = "%B %d, %Y";

#[derive(Debug, StructOpt)]
struct Cli {
    sheet: String,
    #[structopt(parse(from_os_str))]
    path: std::path::PathBuf,
    output: Option<String>,
}

fn main() -> anyhow::Result<()> {
    let env = Env::default().filter_or("RUST_LOG", "info");
    env_logger::init_from_env(env);

    trace!("main start");

    let mut args = Cli::from_args();
    trace!("parsed args {:?}", args);

    let workbook = &mut calamine::open_workbook_auto(args.path.clone())
        .with_context(|| format!("could not open excel workbook at `{:?}`", args.path.clone()))?;
    trace!("opened excel workbook {:?}", args.path);

    let mut date_range =
        excel::timecards::parse_date_range(workbook, &args.sheet).with_context(|| {
            format!(
                "failed to parse dates from workbook sheet `{}`",
                &args.sheet
            )
        })?;
    date_range = date_range_correct_confirmation(&args, &mut date_range)?.clone();

    let employees_vec =
        excel::timecards::parse_worksheet(workbook, &args.sheet, date_range.clone()).with_context(
            || {
                format!(
                    "failed to parse employee data from workbook sheet `{}`",
                    &args.sheet
                )
            },
        )?;

    if args.output.is_none() {
        args.output = Some("output.xlsx".to_string());
    }

    info!("Total employees: {}", employees_vec.len());
    for e in employees_vec.clone() {
        debug!("{:?}", e.id);
        for shift in e.hours {
            debug!("{:?}", shift);
        }
    }

    let filename = args
        .output
        .with_context(|| format!("output filename was specified but is blank!"))?;
    let workbook = xlsxwriter::Workbook::new(&filename);
    excel::sage::generate(workbook, &args.sheet, employees_vec, date_range)
        .with_context(|| "an error occurred while trying try generate spreadsheet")?;

    Ok(())
}

fn date_range_correct_confirmation<'a>(
    args: &'a Cli,
    date_range: &'a mut DateColumnRange,
) -> anyhow::Result<&'a DateColumnRange> {
    let (ref mut start, ref mut end) = date_range.range().with_context(|| {
        format!(
            "failed to parse dates from workbook sheet `{}`",
            &args.sheet
        )
    })?;

    for date in date_range.clone() {
        trace!("{}", date.format(DATE_FORMAT));
    }
    println!(
        "From {} to {}",
        start.format(DATE_FORMAT),
        end.format(DATE_FORMAT)
    );
    let is_dates_correct = dialoguer::Confirm::new()
        .with_prompt("Do these dates look correct?")
        .interact()?;
    if !is_dates_correct {
        println!(
            "Please enter a date in the format {}",
            chrono::Utc::now().format(DATE_FORMAT)
        );

        let start_input: String = dialoguer::Input::new()
            .with_prompt("What is the start date?")
            .interact_text()?;
        let start_naive = chrono::NaiveDate::parse_from_str(&start_input, DATE_FORMAT)
            .with_context(|| format!("please enter date in the format: January 01, 2021"))?;
        let start_date = chrono::Date::<chrono::Utc>::from_utc(start_naive, chrono::Utc);
        date_range.start = Some(start_date);
    }

    let are_columns_correct = dialoguer::Confirm::new()
        .with_prompt(format!(
            "Do the dates range from columns `{}` to `{}`",
            to_column_letter(date_range.head as i32),
            to_column_letter(date_range.tail as i32)
        ))
        .interact()?;
    if !are_columns_correct {
        let start_column_letter: String = dialoguer::Input::new()
            .with_prompt("What column does the start date appear on?")
            .interact_text()?;
        date_range.head = from_column_letter(start_column_letter) as usize;
        let end_column_letter: String = dialoguer::Input::new()
            .with_prompt("What column does the end date appear on?")
            .interact_text()?;
        date_range.tail = from_column_letter(end_column_letter) as usize;
    }

    date_range.end =
        Some(date_range.start.unwrap() + chrono::Duration::days(date_range.len() as i64));

    if !is_dates_correct || !are_columns_correct {
        return date_range_correct_confirmation(args, date_range);
    }

    Ok(date_range)
}
