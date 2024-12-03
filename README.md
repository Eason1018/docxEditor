# DocxEditor

**DocxEditor** is a Python-based program designed to manipulate `.docx` files programmatically. It provides functionality to add, delete, and populate table rows, as well as insert signatures and convert the document to a PDF format.

---

## Features

- **Analyze `.docx` File**: Analyze and display the structure of a `.docx` file, including tables and their rows.
- **Add/Delete Rows in Tables**: Modify the content of tables by adding or deleting rows while maintaining formatting.
- **Populate Tables with CSV Data**: Fill tables in the `.docx` file using data from a CSV file.
- **Insert Signatures**: Add image-based signatures to specific cells in tables, adjusting row heights dynamically.
- **Convert `.docx` to PDF**: Automatically convert the modified `.docx` file into a `.pdf` using Microsoft Word.

---

## Requirements

### Python Libraries

- Python 3.7+
- `python-docx`
- `pypandoc`
- `win32com.client` (requires Microsoft Word)
- `Pillow` (for image processing)

### Software Dependencies

- **Microsoft Word**: Required for `.docx` to `.pdf` conversion.
- **Windows OS**: For `win32com.client`.

---

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/your-repo-name/docxeditor.git
   cd docxeditor ```
2. Set up a virtual environment (optional but recommended):

   ```bash
   Copy code
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate 
3. Install dependencies:

   ```bash
   Copy code
   pip install -r requirements.txt
   Ensure Microsoft Word is installed on your system.

4. Usage
   ```bash
   Place your .docx file and optional data.csv file in the same directory as the script.

5. Run the script:

   ```bash
   Copy code
   python main.py

## Program Workflow
- Analyze the Document: Displays the structure of the .docx file, including tables and their rows.

- Modify Tables: You can add or delete rows in specific tables by entering relevant indices and data when prompted.

- Populate Tables: Data from a CSV file can be used to populate specific tables in the document.

- Add Signatures: Images are inserted into specific table cells, with row heights adjusted to fit the signature dimensions.

- Convert to PDF: The modified .docx file is automatically converted to a .pdf and opened.

## Example
### CSV File Format
For populating tables, the CSV should follow this format:

| Serial No | Tenderer Name          | Tenderer Notified On |
|-----------|-------------------------|-----------------------|
| L7        | Hong Kong Lifts Ltd.    | 2024-11-30           |
| L11       | 三菱電梯香港有限公司        | 2024-12-01           |
| L3        | OKOK 電梯香港有限公司      | 2024-12-02           |
| C5        | BESTBEST 電梯香港有限公司 | 2024-12-03           |

## Signatures
To add a signature, provide the path to the image file. The program dynamically adjusts the row height to display the full signature.

## Troubleshooting
Error: No module named 'win32com'
Ensure pywin32 is installed and Microsoft Word is installed on your machine.

## Signatures not fully visible
Ensure row height adjustment is enabled in the script.

## Conversion Issues
Make sure Microsoft Word is installed and added to the system's PATH.

## Contributing
Contributions are welcome! Feel free to submit pull requests for enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgements
python-docx

Pillow

PyWin32

