use calamine::{Data, Reader, Xlsx, open_workbook};

use std::fs::File;
use std::io::BufReader;

trait MyStage<T, U> {
    fn run(&self, x: T) -> U;
}

impl<T, U, F> MyStage<T, U> for F
where
    F: Fn(T) -> U,
{
    fn run(&self, x: T) -> U {
        self(x)
    }
}

fn get_open(x: &str) -> Result<Xlsx<BufReader<File>>, calamine::Error> {
    open_workbook(x).map_err(calamine::Error::from)
}

fn main() {
    // 1. read path
    let a1 = "/mnt/c/Users/stewartli/Desktop/tbd/BF164 (31 July 2026)/BF164v FP&A isca-ISCA 31 Jul 2026 PASSWORD.xlsx";
    let _ = (get_open).run(a1);

    let mut wb = match get_open(a1) {
        Ok(x) => x,
        Err(calamine::Error::Xlsx(calamine::XlsxError::Password)) => {
            println!("ask your password");
            return;
        }
        Err(x) => {
            println!("error: {x}");
            return;
        }
    };

    // 2. sheet names
    println!("\x1b[33mSheets\x1b[0m");
    wb.sheet_names().iter().for_each(|x| {
        println!("{x:<width$}", width = 50);
    });

    // 3. sheet formula
    let wb1 = wb.with_header_row(calamine::HeaderRow::Row(0));
    let fm = wb1.worksheet_formula("Assumptions").unwrap();
    let rg = wb1.worksheet_range("Assumptions").unwrap();
    let (srn, scn) = fm.start().unwrap_or((0, 0));

    let n_fm = fm
        .rows()
        .flat_map(|r| r.iter().filter(|f| !f.is_empty()))
        .count();

    println!("\x1b[33mFormulas: {n_fm}\x1b[0m");
    println!("\x1b[32mSize: {:?}\x1b[0m", rg.get_size());

    for (r, row) in fm.rows().enumerate() {
        for (c, formula) in row.iter().enumerate() {
            if formula.is_empty() {
                continue;
            }
            let row = srn + r as u32;
            let col = scn + c as u32;
            let data = rg
                .get_value((row, col))
                .map(|v| v.to_string())
                .unwrap_or_default();
            println!("[{:<3},{:<3}] {:<20} {}", row, col, data, formula);
        }
    }

    // 4. get value
    let out = rg
        .used_cells()
        .filter(|(_, _, x)| **x == Data::String("Worst -20%, Base -10%, Best +0%".into()))
        .map(|(r, c, _)| (r as u32, c as u32))
        .collect::<Vec<(u32, u32)>>();

    println!("Search: {:?}", out[0]);
}
