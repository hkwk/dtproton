use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use regex::Regex;

fn cell_ref(col_1_based: usize, row_1_based: usize) -> String {
    fn col_to_name(mut col: usize) -> String {
        // 1 -> A, 26 -> Z, 27 -> AA
        let mut name = String::new();
        while col > 0 {
            let rem = (col - 1) % 26;
            name.push((b'A' + rem as u8) as char);
            col = (col - 1) / 26;
        }
        name.chars().rev().collect()
    }

    format!("{}{}", col_to_name(col_1_based), row_1_based)
}

fn datatype_to_string(v: &Data) -> String {
    match v {
        Data::Empty => String::new(),
        Data::String(s) => s.clone(),
        Data::Float(f) => {
            if f.fract() == 0.0 {
                format!("{:.0}", f)
            } else {
                f.to_string()
            }
        }
        Data::Int(i) => i.to_string(),
        Data::Bool(b) => b.to_string(),
        Data::DateTime(f) => f.to_string(),
        Data::DateTimeIso(s) => s.clone(),
        Data::DurationIso(s) => s.clone(),
        Data::Error(e) => format!("{e:?}"),
    }
}

fn processed_output_path(input: &Path) -> PathBuf {
    let file_name = input
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "output.xlsx".to_string());
    PathBuf::from(format!("processed_{file_name}"))
}

fn process_excel(path: &Path) -> Result<Option<PathBuf>> {
    // Use umya-spreadsheet to determine the active sheet (to match excelize behavior),
    // and later to write the modified workbook back out.
    let mut book = umya_spreadsheet::reader::xlsx::read(path)
        .with_context(|| format!("无法读取文件: {}", path.display()))?;
    let active_sheet_index = *book.get_workbook_view().get_active_tab() as usize;

    let mut workbook = open_workbook_auto(path).with_context(|| format!("无法打开文件: {}", path.display()))?;
    let sheet_names = workbook.sheet_names();
    let sheet_name = sheet_names
        .get(active_sheet_index)
        .or_else(|| sheet_names.get(0))
        .cloned()
        .ok_or_else(|| anyhow!("工作簿中没有工作表"))?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .with_context(|| format!("无法读取工作表: {sheet_name}"))?;

    // A3 -> row=3, col=A => (2,0) in 0-based
    let a3 = range
        .get_value((2u32, 0u32))
        .map(datatype_to_string)
        .unwrap_or_default();

    if a3.trim() != "离子色谱" {
        println!("A3 单元格不是“离子色谱”，无需处理。");
        return Ok(None);
    }

    let (height, width) = range.get_size();
    if height < 6 {
        println!("表格行数不足6行，无需处理第6行及以后的数据。");
        return Ok(None);
    }

    let re = Regex::new(r"\((RM|C)\)").expect("valid regex");

    // Collect cells to clear (1-based coordinates for Excel refs)
    let mut to_clear: Vec<String> = Vec::new();
    for row0 in 5..height {
        for col0 in 0..width {
            let value = range
                .get_value((row0 as u32, col0 as u32))
                .map(datatype_to_string)
                .unwrap_or_default();
            if !value.is_empty() && re.is_match(&value) {
                to_clear.push(cell_ref(col0 + 1, row0 + 1));
            }
        }
    }

    if to_clear.is_empty() {
        // Still mimic Go behavior: save only if changes? In Go it always SaveAs.
        // We'll still save a copy so behavior matches "processed_" output.
        // (If you prefer skipping when no changes, tell me.)
    }

    // calamine is read-only; use umya-spreadsheet to write the updated workbook.
    let sheet = book.get_active_sheet_mut();

    for addr in to_clear {
        sheet.get_cell_mut(addr.as_str()).set_value("");
    }

    let output_path = processed_output_path(path);
    umya_spreadsheet::writer::xlsx::write(&book, &output_path)
        .with_context(|| format!("无法保存文件: {}", output_path.display()))?;

    Ok(Some(output_path))
}

fn main() {
    if let Err(e) = real_main() {
        eprintln!("处理 Excel 文件时出错: {e:#}");
        std::process::exit(1);
    }
}

fn real_main() -> Result<()> {
    let mut args = std::env::args_os();
    let _exe = args.next();
    let Some(input) = args.next() else {
        println!("请提供文件名作为参数，例如：dtproton 45vocs2.xlsx");
        return Ok(());
    };

    let input_path = PathBuf::from(input);
    let out = process_excel(&input_path)?;
    if let Some(out) = out {
        println!("文件已处理并保存为: {}", out.display());
    }
    Ok(())
}
