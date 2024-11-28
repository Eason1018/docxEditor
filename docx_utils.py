import os
import zipfile
import csv
from lxml import etree
from docx import Document


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

    # Analyze all tables in the document
    for table_index, table in enumerate(doc.tables):
        print(f"\nTable {table_index}:")
        for row_index, row in enumerate(table.rows):
            row_data = [cell.text.strip() for cell in row.cells]
            print(f" Row {row_index}: {row_data}")

def modify_document_xml(document_xml_path, replacements):
    """
    Modifies the document.xml file by replacing specified text.
    """
    # Parse the XML
    parser = etree.XMLParser(ns_clean=True, recover=True)
    tree = etree.parse(document_xml_path, parser)
    root = tree.getroot()

    # Namespaces used in the Word XML
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Replace text in document.xml
    for text_element in root.xpath('.//w:t', namespaces=namespaces):
        if text_element.text and any(old_text in text_element.text for old_text in replacements):
            for old_text, new_text in replacements.items():
                if old_text in text_element.text:
                    text_element.text = text_element.text.replace(old_text, new_text)

    # Save the changes back to the file
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


def populate_table_from_csv(docx_file, output_file, csv_file, table_index=0):
    """
    Populate a Word document table using data from a CSV file.
    """
    # Load the Word document
    doc = Document(docx_file)

    # Select the target table
    table = doc.tables[table_index]

    # Read data from CSV
    with open(csv_file, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        headers = next(reader)  # Skip the header row in the CSV file

        # Populate the table rows
        for i, row_data in enumerate(reader):
            if i + 1 < len(table.rows):  # Skip the header row in the Word table
                row = table.rows[i + 1]
                for j, cell_data in enumerate(row_data):
                    row.cells[j].text = cell_data
            else:
                # Add a new row if necessary
                new_row = table.add_row()
                for j, cell_data in enumerate(row_data):
                    new_row.cells[j].text = cell_data

    # Save the modified document
    doc.save(output_file)

    def populate_table_0(table):
        """
        Populate specific fields in Table 0 with hardcoded data.
        """
        # Hardcoded data for Table 0
        table_0_data = {
            "From:": "Ng, Wai Ming, Rock / 吳偉明 / 88888",
            "To:": "Tendering Committee",
            "Name of Property:": "AP- Apec Plaza",
            "The Works:": "電機控制系統問題扶手鬆動滑輪/滑道磨損安全營示系統失靈：運行並確認問題維修過程需要專業技術人員進行,確保電梯/扶手電梯安全可靠運行",
            "Tender Ref.:": "SHKP1234-001"
        }

        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip() in table_0_data:  # Check if the text matches a key
                    key = cell.text.strip()
                    cell.text = table_0_data[key]