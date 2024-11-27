import csv
import os
from pdfrw import PdfReader, PdfWriter, PdfDict

# Create output directory if it doesn't exist
output_directory = 'output'
os.makedirs(output_directory, exist_ok=True)

# Function to replace fields in the PDF
def fill_pdf(template_path, output_path, data):
    template = PdfReader(template_path)
    for page in template.pages:
        annotations = page['/Annots']
        if annotations:
            for annotation in annotations:
                field_name = annotation.get('/T')  # Safely get the field name
                if field_name:  # Check if field_name is not None
                    field_name = field_name[1:-1]  # Remove parentheses
                    if field_name in data:
                        annotation.update(
                            PdfDict(V='{}'.format(data[field_name]))
                        )
    PdfWriter().write(output_path, template)

# Paths to your files
csv_file_path = 'data.csv'  # Your CSV file
pdf_template_path = 'template.pdf'  # Your PDF template

# Read the CSV and create PDFs
with open(csv_file_path, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile, delimiter=';')
    for row in reader:
        # Clean the KD field for the filename
        kd_filename = row['KD'].replace('/', '').replace('_', '')  # Remove special characters
        output_pdf_path = os.path.join(output_directory, f'{kd_filename}.pdf')  # Create output PDF path
        fill_pdf(pdf_template_path, output_pdf_path, row)  # Fill PDF and save

print("PDFs generated successfully in the 'output' directory.")
