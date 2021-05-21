extern crate chrono;

#[derive(Clone)]
pub struct Employee {
    pub id: String,
    pub hours: Vec<Shift>,
    pub overtime_schedule: String,
    pub dist_code: String,
    pub exp_account: String,
}

impl Employee {
    pub fn new(id: String) -> Self {
        Self {
            id,
            hours: Vec::new(),
            overtime_schedule: String::new(),
            dist_code: String::new(),
            exp_account: String::new(),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub struct Shift {
    pub col: usize,
    pub duration: chrono::Duration,
    pub date: chrono::Date<chrono::Utc>,
}

impl Shift {
    pub fn sum_of_shift(&self) -> f32 {
        self.duration.num_minutes() as f32 / 60.0
    }
}

pub fn sum_of_hours(shifts: Vec<Shift>) -> f32 {
    let mut sum = 0.0;

    for shift in shifts.iter() {
        sum += shift.sum_of_shift();
    }

    sum
}
