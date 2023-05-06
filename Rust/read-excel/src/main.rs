// fn main() {
//     println!("Hello, world!");
// }

// Bing GPT
use calamine::open_workbook;
use calamine::Result;

fn main() {
    let path = "C:\\Users\\GB675AG\\EY\\Projeto Samsung Order Mgmt - General\\03. Gestão da Rotina\\Automações\\Ferramentas\\BCK\\IM\\01.02.2022\\OUTBOUND IM.xlsx";
    let mut workbook: Result<_> = open_workbook(path);
    let sheet_name = workbook.as_ref().unwrap().sheet_names().get(0).unwrap();
    println!("Sheet name: {}", sheet_name);
}

// Chat GPT
// use calamine::{Reader, Xlsx};

// fn main() {
//     let file_path = "C:\\Users\\GB675AG\\EY\\Projeto Samsung Order Mgmt - General\\03. Gestão da Rotina\\Automações\\Ferramentas\\BCK\\IM\\01.02.2022\\OUTBOUND IM.xlsx";

//     // Open the Excel file
//     let mut excel: Xlsx<_> = Reader::open(file_path).unwrap();

//     // Read the names of worksheets (tabs)
//     let sheet_names = excel.sheet_names().unwrap();

//     // Print the sheet names
//     for name in sheet_names {
//         println!("{}", name);
//     }
// }
