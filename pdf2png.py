from pdf2image import convert_from_path
import os

# Set the input and output directories
input_dir = 'D:\EY\Cerificate Scanner\Cerificates'  # Replace with the path to your input directory
output_dir = 'D:\EY\Cerificate Scanner\Certificate png'  # Replace with the path to your output directory

# Get the list of PDF files in the input directory
pdf_files = [f for f in os.listdir(input_dir) if f.endswith('.pdf')]

for pdf_file in pdf_files:
    pdf_path = os.path.join(input_dir, pdf_file)

    # Convert PDF to images using pdf2image
    images = convert_from_path(pdf_path)

    # Save each page of the PDF as a separate PNG file
    for i, image in enumerate(images):
        output_file = os.path.splitext(pdf_file)[0] + f'_page{i + 1}.png'
        output_path = os.path.join(output_dir, output_file)
        image.save(output_path, 'PNG')

    print(f'Converted {pdf_file} to PNG.')
