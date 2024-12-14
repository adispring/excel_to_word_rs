use calamine::{open_workbook, Reader, Xlsx};
use docx_rs::*;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Read the Excel file
    let mut workbook: Xlsx<_> = match open_workbook("src/input.xlsx") {
        Ok(wb) => wb,
        Err(e) => {
            eprintln!("Error opening workbook: {}", e);
            return Err(Box::new(e));
        }
    };
    let range = workbook
        .worksheet_range("Sheet1")
        .ok_or("Cannot find sheet")??;

    // Create a new Word document
    let mut doc = Docx::new();

    for (i, row) in range.rows().enumerate() {
        if i == 0 {
            continue; // Skip header row
        }

        let level1 = row.get(0).map(|c| c.to_string()).unwrap_or_default();
        let level2 = row.get(1).map(|c| c.to_string()).unwrap_or_default();
        let level3 = row.get(2).map(|c| c.to_string()).unwrap_or_default();
        let content1 = row.get(3).map(|c| c.to_string()).unwrap_or_default();
        let content3 = row.get(5).map(|c| c.to_string()).unwrap_or_default();

        if !level1.is_empty() {
            doc = doc.add_paragraph(
                Paragraph::new().add_run(Run::new().add_text(level1).bold().size(48).heading1()),
            );
        }

        if !level2.is_empty() {
            doc = doc.add_paragraph(
                Paragraph::new().add_run(Run::new().add_text(level2).bold().size(40).heading2()),
            );
        }

        if !level3.is_empty() {
            doc = doc.add_paragraph(
                Paragraph::new().add_run(Run::new().add_text(level3).bold().size(36).heading3()),
            );
        }

        let concatenated_content = [content1, content3]
            .iter()
            .filter(|&s| !s.is_empty())
            .cloned()
            .collect::<Vec<_>>()
            .join("; ");

        if !concatenated_content.is_empty() {
            doc = doc.add_paragraph(
                Paragraph::new().add_run(Run::new().add_text(concatenated_content).size(24)),
            );
        }
    }

    // Save the document
    let mut file = File::create("output/output.docx")?;
    doc.build().pack(&mut file)?;

    println!("Document saved successfully to output/output.docx");

    Ok(())
}
