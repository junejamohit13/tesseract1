import os
from docx import Document
import openpyxl
from openpyxl.styles import Font
import re

def extract_tables_and_headings(docx_path):
    doc = Document(docx_path)
    tables_and_headings = []
    paragraphs = doc.paragraphs
    tables = doc.tables
    num_paragraphs = len(paragraphs)
    table_index = 0

    for i in range(num_paragraphs):
        paragraph = paragraphs[i]
        # Check if the paragraph matches the "Table X:" pattern
        if re.match(r'^Table \d+:', paragraph.text.strip()):
            # Ensure there is a table following this heading
            if table_index < len(tables):
                # Check if the next element is a table
                if paragraph._element.getnext() is not None and paragraph._element.getnext().tag.endswith('tbl'):
                    # Assign the heading to the table
                    table = tables[table_index]
                    tables_and_headings.append((paragraph.text.strip(), table))
                    table_index += 1

    return tables_and_headings

def create_excel_from_directory(directory_path, excel_path):
    wb = openpyxl.Workbook()
    summary_sheet = wb.active
    summary_sheet.title = "Summary"
    
    # Creating headers for summary sheet
    headers = ["Document Name", "Table Heading", "Link to Table"]
    summary_sheet.append(headers)
    for col in range(1, len(headers) + 1):
        summary_sheet.cell(row=1, column=col).font = Font(bold=True)

    for filename in os.listdir(directory_path):
        if filename.endswith('.docx'):
            docx_path = os.path.join(directory_path, filename)
            doc_name = os.path.splitext(filename)[0]
            tables_and_headings = extract_tables_and_headings(docx_path)
            
            for idx, (heading, table) in enumerate(tables_and_headings, start=1):
                sheet_name = f"{doc_name}_Table_{idx}"
                table_sheet = wb.create_sheet(title=sheet_name)
                
                # Write table data to the table_sheet
                for i, row in enumerate(table.rows, start=1):
                    for j, cell in enumerate(row.cells, start=1):
                        table_sheet.cell(row=i, column=j).value = cell.text

                # Adding data to the summary sheet
                summary_sheet.append([doc_name, heading, f"=HYPERLINK(\"#{sheet_name}!A1\", \"{sheet_name}\")"])
                link_cell = summary_sheet.cell(row=summary_sheet.max_row, column=3)
                # Set the font to blue and underlined
                link_cell.font = Font(color="0000FF", underline="single")

    wb.save(excel_path)

create_excel_from_directory(directory_path, excel_path)
