import os
import zipfile
import csv

from docx.oxml import OxmlElement
from lxml import etree
from docx import Document
import pypandoc
import win32com.client

def extract_docx(docx_path, extract_dir):
    """
    Extracts a .docx file into a specified directory.
    """
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)


def analyze_document(docx_file):
    """
    Analyze the structure of the Word document (e.g., tables) and log it.
    """
    print(f"Analyzing document: {docx_file}")
    doc = Document(docx_file)
    for table_index, table in enumerate(doc.tables):
        print(f"\nTable {table_index}:")
        for row_index, row in enumerate(table.rows):
            row_data = [cell.text.strip() for cell in row.cells]
            print(f" Row {row_index}: {row_data}")


def modify_document_xml(document_xml_path, replacements):
    """
    Modifies the document.xml file by replacing specified text.
    """
    parser = etree.XMLParser(ns_clean=True, recover=True)
    tree = etree.parse(document_xml_path, parser)
    root = tree.getroot()
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for text_element in root.xpath('.//w:t', namespaces=namespaces):
        if text_element.text and any(old_text in text_element.text for old_text in replacements):
            for old_text, new_text in replacements.items():
                if old_text in text_element.text:
                    text_element.text = text_element.text.replace(old_text, new_text)

    tree.write(document_xml_path, xml_declaration=True, encoding='UTF-8', standalone="yes")


def repack_docx(extract_dir, output_docx_path):
    """
    Repackages the extracted files back into a .docx file.
    """
    with zipfile.ZipFile(output_docx_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root_dir, dirs, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, extract_dir)
                zipf.write(file_path, arcname)


def populate_table_0(table):
    """
    Populate specific fields in Table 0 with hardcoded data.
    """
    table_0_data = {
        "From:": "Ng, Wai Ming, Rock / 吳偉明 / 88888",
        "To:": "Tendering Committee",
        "Name of Property:": "AP- Apec Plaza",
        "The Works:": "電機控制系統問題扶手鬆動滑輪/滑道磨損安全營示系統失靈：運行並確認問題維修過程需要專業技術人員進行,確保電梯/扶手電梯安全可靠運行",
        "Tender Ref.:": "SHKP1234-001"
    }

    for row in table.rows:
        cells = row.cells
        if len(cells) >= 2:
            for i, cell in enumerate(cells):
                if cell.text.strip() in table_0_data:
                    key = cell.text.strip()
                    cells[i + 1].text = table_0_data[key]


def populate_table_from_csv(doc, csv_file, table_index=1):
    """
    Populate Table 1 using data from a CSV file.
    """
    table = doc.tables[table_index]

    with open(csv_file, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        headers = next(reader)

        row_idx = 2
        for row_data in reader:
            while row_idx < len(table.rows):
                table_row = table.rows[row_idx]
                if all(not cell.text.strip() for cell in table_row.cells):
                    row_idx += 1
                    continue

                for csv_col, cell_data in zip(headers, row_data):
                    for cell in table_row.cells:
                        if csv_col.strip() == cell.text.strip():
                            cell.text = cell_data.strip()
                row_idx += 1
                break


def add_row_to_table(doc, table_index, row_data):
    """
    Add a new row to the specified table with the given data and maintain formatting.
    """
    table = doc.tables[table_index]

    # Use the last row as a template for formatting
    template_row = table.rows[-1]

    # Add a new row
    new_row = table.add_row()

    # Populate the new row with the provided data
    for i, cell_data in enumerate(row_data):
        if i < len(new_row.cells):  # Ensure the data fits in the table
            new_cell = new_row.cells[i]
            new_cell.text = cell_data

            # Copy formatting from the template row
            template_cell = template_row.cells[i] if i < len(template_row.cells) else None
            if template_cell:
                _copy_cell_formatting(template_cell, new_cell)

    print(f"Row added to Table {table_index} with data: {row_data}")


def _copy_cell_formatting(template_cell, target_cell):
    """
    Copy the text formatting from the template cell to the target cell.
    """
    # Copy paragraph alignment
    if template_cell.paragraphs and target_cell.paragraphs:
        target_cell.paragraphs[0].alignment = template_cell.paragraphs[0].alignment
    template_paragraph = template_cell.paragraphs[0]
    target_paragraph = target_cell.paragraphs[0]

    # Copy font settings
    if template_paragraph.runs and target_paragraph.runs:
        template_font = template_paragraph.runs[0].font
        target_font = target_paragraph.runs[0].font

        # Copy font size, name, and other attributes
        if template_font.size:
            target_font.size = template_font.size
        if template_font.name:
            target_font.name = template_font.name
        target_font.bold = template_font.bold
        target_font.italic = template_font.italic
        target_font.underline = template_font.underline

    # Copy alignment (if needed)
    target_paragraph.alignment = template_paragraph.alignment



def delete_row_from_table(doc, table_index, row_number):
    """
    Delete a specific row from the specified table.
    """
    table = doc.tables[table_index]

    if row_number < 0 or row_number >= len(table.rows):
        print(f"Invalid row number: {row_number}")
        return

    # Access the row to delete
    row_to_delete = table.rows[row_number]

    # Remove the row from the table
    tbl = table._tbl
    tbl.remove(row_to_delete._tr)

    print(f"Row {row_number} deleted from Table {table_index}.")


def convert_to_pdf(input_docx, output_pdf):
    """
    Convert a .docx file to .pdf using win32com.client (requires Microsoft Word).
    """
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(input_docx)
        doc.SaveAs(output_pdf, FileFormat=17)  # 17 is the constant for wdFormatPDF
        doc.Close()
        word.Quit()
        print(f"PDF file created: {output_pdf}")
    except Exception as e:
        print(f"Error during PDF conversion: {e}")
