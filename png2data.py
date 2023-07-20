import pytesseract
from PIL import Image
import re
import os
import pandas as pd

# Set the Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_information_from_image(image_path):
    im = Image.open(image_path)
    text = pytesseract.image_to_string(im, lang='eng')

    name_pattern = r"(.*?)\s+(?:S/o|W/o|D/o)\s+(.*?)\n"
    course_pattern = r"for the role of\n(.*?)\s+\((.*?)\)"
    level_pattern = r"National Skills Qualifications Framework Level-(\d+)"

    candidate_match = re.search(name_pattern, text)
    candidate_name = candidate_match.group(1) if candidate_match else None
    guardian_name = candidate_match.group(2) if candidate_match else None

    course_match = re.search(course_pattern, text)
    course_name = course_match.group(1) if course_match else None
    course_code = course_match.group(2) if course_match else None

    level_match = re.search(level_pattern, text)
    skill_level = level_match.group(1) if level_match else None

    return {
        "Candidate Name": candidate_name,
        "Guardian Name": guardian_name,
        "Course Name": course_name,
        "Course Code": course_code,
        "Level": skill_level
    }

def process_png_files(folder_path, output_file):
    data = []

    # Process each PNG file in the folder
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith('.png'):
            file_path = os.path.join(folder_path, file_name)
            extracted_info = extract_information_from_image(file_path)
            extracted_info['File Name'] = file_name
            data.append(extracted_info)

    # Create a DataFrame and save it to an Excel file
    df = pd.DataFrame(data)

    # Move 'File Name' to the first column
    df = df[['File Name'] + [col for col in df.columns if col != 'File Name']]

    df.to_excel(output_file, index=False)

folder_path = "D:\EY\Cerificate Scanner\Certificate png"
output_file = "D:\EY\Cerificate Scanner\output42.xlsx"

process_png_files(folder_path, output_file)
print("Extraction completed and data saved to", output_file)
