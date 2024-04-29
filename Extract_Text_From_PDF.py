import fitz  # PyMuPDF
import re
import csv

# PDF file path
pdf_path = r'D:\invoice.pdf'
# Path to the CSV file
output_path = r"D:\invoice.csv"

try:
    # Open the PDF file
    with fitz.open(pdf_path) as pdf_document:
        # Initialize variables to store extracted text
        pdf_text = ""

        # Iterate through each page of the PDF
        for page_num in range(len(pdf_document)):
            # Get the page text
            page_text = pdf_document[page_num].get_text()
            # Append the page text to the overall PDF text
            pdf_text += page_text

        # Regular expressions to extract invoice number and date
        invoice_number_match = re.search(r'Invoice Number\s+([\w\d-]+)', pdf_text)
        invoice_date_match = re.search(r'Invoice Date\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})', pdf_text)

        # Check if both invoice number and date were found
        if invoice_number_match and invoice_date_match:
            invoice_number = invoice_number_match.group(1)
            invoice_date = invoice_date_match.group(1)
            print("Invoice Number:", invoice_number)
            print("Invoice Date:", invoice_date)

            # Write data to CSV file
            with open(output_path, "w", newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow(["Invoice Number", "Invoice Date"])
                csv_writer.writerow([invoice_number, invoice_date])

            print("Invoice number and date written to invoice.csv successfully!")
        else:
            print("Failed to extract invoice number and date from the PDF.")

except Exception as e:
    print(f"An error occurred: {e}")
