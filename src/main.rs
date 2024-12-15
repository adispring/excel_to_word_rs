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
    let mut file = File::create("output/wrong.docx")?;
    let range = workbook
        .worksheet_range("Sheet1")
        .ok_or("Cannot find sheet")??;

    let style1 = Style::new("Heading1", StyleType::Paragraph).name("Heading 1");
    let style2 = Style::new("Heading2", StyleType::Paragraph).name("Heading 2");
    let style3 = Style::new("Heading3", StyleType::Paragraph).name("Heading 3");

    // Create a new Word document
    let mut doc = Docx::new()
        .add_style(style1)
        .add_style(style2)
        .add_style(style3);

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
                Paragraph::new()
                    .add_run(Run::new().add_text(level1).bold().size(48))
                    .style("Heading1"), // .page_break_before(true),
            );
        }

        if !level2.is_empty() {
            doc = doc.add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text(level2).bold().size(40))
                    .style("Heading2"),
            );
        }

        if !level3.is_empty() {
            doc = doc.add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text(level3).bold().size(36))
                    .style("Heading3"),
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

    doc.build().pack(&mut file)?;

    println!("Document saved successfully to output/output.docx");

    Ok(())
}
