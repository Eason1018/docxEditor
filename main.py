import os
from docx_utils import *

# File paths
INPUT_DOCX = "input.docx"  # Replace with your actual .docx file path
OUTPUT_DOCX = "output.docx"  # Path for the modified .docx file
INPUT_CSV = "data.csv"  # Path to the CSV file

# Ensure input files exist
if not os.path.exists(INPUT_DOCX):
    print(f"Error: Input file '{INPUT_DOCX}' does not exist.")
    exit(1)

if not os.path.exists(INPUT_CSV):
    print(f"Error: Input file '{INPUT_CSV}' does not exist.")
    exit(1)

# Delete the existing output file if it exists
if os.path.exists(OUTPUT_DOCX):
    print(f"Deleting existing output file: {OUTPUT_DOCX}")
    os.remove(OUTPUT_DOCX)

# Step 1: Analyze the input document structure
print("\nSTEP 1: Analyzing Input Document")
analyze_document(INPUT_DOCX)

# Load the Word document
doc = Document(INPUT_DOCX)

# # Step 2: Populate Table 0 with hardcoded data
# print("\nSTEP 2: Populating Table 0 with Hardcoded Data...")
# try:
#     populate_table_0(doc.tables[0])
#     print("Table 0 populated successfully.")
# except Exception as e:
#     print(f"Error during Table 0 population: {e}")
#
# # Step 3: Populate Table 1 using CSV data
# print("\nSTEP 3: Populating Table 1 from CSV...")
# try:
#     populate_table_from_csv(doc, INPUT_CSV, table_index=1)
#     print("Table 1 populated successfully.")
# except Exception as e:
#     print(f"Error during Table 1 population: {e}")

# Step 4: Modify the document (Add/Delete rows)
modify_choice = input("Do you want to modify the document (add/delete rows)? (yes/no): ").strip().lower()

if modify_choice in ["yes", "y"]:
    while True:
        # Console interaction for modifications
        print("Choose an action: ")
        print("1. Add a row")
        print("2. Delete a row")
        print("3. Add a signature to a cell")
        print("4. Exit modifications")
        choice = input("Enter 1, 2, 3, or 4: ").strip()

        if choice == "1":
            # Add a row
            try:
                table_index = int(input("Enter the table index (0, 1, etc.): "))
                row_data = input("Enter row data as comma-separated values: ").split(",")
                add_row_to_table(doc, table_index, row_data)
                print(f"Row added to Table {table_index}.")
            except Exception as e:
                print(f"Error adding row: {e}")

        elif choice == "2":
            # Delete a row
            try:
                table_index = int(input("Enter the table index (0, 1, etc.): "))
                row_number = int(input("Enter the row number to delete (starting from 0): "))
                delete_row_from_table(doc, table_index, row_number)
                print(f"Row {row_number} deleted from Table {table_index}.")
            except Exception as e:
                print(f"Error deleting row: {e}")


        elif choice == "3":

            # Add a signature to a cell

            try:

                table_index = int(input("Enter the table index (0, 1, etc.): "))
                row_index = int(input("Enter the row index (starting from 0): "))
                column_index = int(input("Enter the column index (starting from 0): "))
                image_path = input("Enter the path to the signature image: ")

                table = doc.tables[table_index]
                cell = table.rows[row_index].cells[column_index]
                add_signature_to_cell(cell, image_path)
                print(f"Signature added to Table {table_index}, Row {row_index}, Column {column_index}.")
            except Exception as e:
                print(f"Error adding signature: {e}")

        elif choice == "4":
            print("Exiting modifications...")
            break

        else:
            print("Invalid choice! Please enter 1, 2, 3, or 4.")


# Save the modified document
doc.save(OUTPUT_DOCX)
print(f"Document saved as: {OUTPUT_DOCX}")

# Step 5: Convert to PDF
OUTPUT_PDF = OUTPUT_DOCX.replace('.docx', '.pdf')
try:
    convert_to_pdf(os.path.abspath(OUTPUT_DOCX), os.path.abspath(OUTPUT_PDF))
    # Automatically open the generated PDF file
    if os.path.exists(OUTPUT_PDF):
        os.startfile(OUTPUT_PDF)
        print(f"PDF file opened: {OUTPUT_PDF}")
except Exception as e:
    print(f"Error during PDF conversion: {e}")

print(f"\nProcess completed. Modified document saved as: {OUTPUT_DOCX}")
print(f"PDF output saved as: {OUTPUT_PDF}")
