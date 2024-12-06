import csv
import os
import subprocess

# Path to pdftk executable
pdftk_path = "pdftk.exe"

# Create output directories
output_directory = "Output"
new_directory = "New"  # Create New folder in the root directory
os.makedirs(output_directory, exist_ok=True)
os.makedirs(new_directory, exist_ok=True)

# Paths to your files
csv_file_path = "data.csv"  # Your CSV file
pdf_template_path = "template.pdf"  # Your PDF template

# Function to encode strings for pdftk
def encode_for_pdftk(value):
    # Encode as Latin-1 (ISO-8859-1) compatible with pdftk
    try:
        return value.encode("latin-1", "replace").decode("latin-1")
    except Exception:
        return value  # Return the original if encoding fails

# Function to generate unique filenames for repeated KDs
def generate_unique_filename(base_name, directory):
    counter = 1
    unique_name = f"{base_name}.pdf"
    while os.path.exists(os.path.join(directory, unique_name)):
        unique_name = f"{base_name}_{counter}.pdf"
        counter += 1
    return unique_name

# Read the CSV and create flattened PDFs
with open(csv_file_path, newline="", encoding="utf-8-sig") as csvfile:  # Handle UTF-8 with BOM
    reader = csv.DictReader(csvfile, delimiter=";")
    for row in reader:
        # Generate base filename from KD field
        base_filename = row["KD"].replace("/", "").replace("_", "").strip()
        output_pdf_path = os.path.join(output_directory, base_filename + ".pdf")

        # Check if the file already exists in the Output folder
        if os.path.exists(output_pdf_path):
            print(f"File '{output_pdf_path}' already exists in Output. Skipping...")
            continue

        # Generate a unique filename for New folder
        unique_filename = generate_unique_filename(base_filename, new_directory)
        new_pdf_path = os.path.join(new_directory, unique_filename)

        # Prepare the temporary FDF file for form data
        fdf_content = "%FDF-1.2\n%\xE2\xE3\xCF\xD3\n1 0 obj\n<< /FDF << /Fields ["
        for key, value in row.items():
            encoded_value = encode_for_pdftk(value)
            fdf_content += f"<< /T ({key}) /V ({encoded_value}) >>\n"
        fdf_content += "] >> >>\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF"

        # Write the FDF content to a temporary file
        fdf_file = "temp.fdf"
        with open(fdf_file, "w", encoding="latin-1") as f:  # Write FDF in Latin-1
            f.write(fdf_content)

        # Use pdftk to fill the form and flatten the PDF
        command = [
            pdftk_path,
            pdf_template_path,
            "fill_form",
            fdf_file,
            "output",
            new_pdf_path,
            "flatten",
            "compress",
        ]
        subprocess.run(command, check=True)

        # Remove the temporary FDF file
        os.remove(fdf_file)

print(f"PDFs generated successfully in the 'New' folder.")
