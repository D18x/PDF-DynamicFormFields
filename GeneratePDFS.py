import csv
import os
import subprocess

# Path to pdftk executable
pdftk_path = "pdftk.exe"

# Create output directory if it doesn't exist
output_directory = "output"
os.makedirs(output_directory, exist_ok=True)

# Paths to your files
csv_file_path = "data.csv"  # Your CSV file
pdf_template_path = "template.pdf"  # Your PDF template

# Read the CSV and create flattened PDFs
with open(csv_file_path, newline="", encoding="utf-8") as csvfile:
    reader = csv.DictReader(csvfile, delimiter=";")
    for row in reader:
        # Prepare the temporary FDF file for form data
        fdf_content = "%FDF-1.2\n%\xE2\xE3\xCF\xD3\n1 0 obj\n<< /FDF << /Fields ["
        for key, value in row.items():
            fdf_content += f"<< /T ({key}) /V ({value}) >>\n"
        fdf_content += "] >> >>\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF"

        # Write the FDF content to a temporary file
        fdf_file = "temp.fdf"
        with open(fdf_file, "w", encoding="utf-8") as f:
            f.write(fdf_content)

        # Generate output PDF filename
        kd_filename = row["KD"].replace("/", "").replace("_", "").strip()
        output_pdf_path = os.path.join(output_directory, f"{kd_filename}.pdf")

        # Use pdftk to fill the form and flatten the PDF
        command = [
            pdftk_path,
            pdf_template_path,
            "fill_form",
            fdf_file,
            "output",
            output_pdf_path,
            "flatten",
            "compress",
        ]
        subprocess.run(command, check=True)

        # Remove the temporary FDF file
        os.remove(fdf_file)

print("PDFs generated successfully in the 'output' directory.")
