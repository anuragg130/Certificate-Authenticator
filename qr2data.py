import pytesseract
from PIL import Image
import re
import os
import pandas as pd
from qreader import QReader
import cv2
import pyzbar
from pyzbar import *

# Set the Tesseract OCR path
#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# def extract_information_from_image(image_path):
#     im = Image.open(image_path)
#     text = pytesseract.image_to_string(im, lang='eng')
#
#     name_pattern = r"(.*?)\s+(?:S/o|W/o|D/o)\s+(.*?)\n"
#     course_pattern = r"for the role of\n(.*?)\s+\((.*?)\)"
#     level_pattern = r"National Skills Qualifications Framework Level-(\d+)"
#
#     candidate_match = re.search(name_pattern, text)
#     candidate_name = candidate_match.group(1) if candidate_match else None
#     guardian_name = candidate_match.group(2) if candidate_match else None
#
#     course_match = re.search(course_pattern, text)
#     course_name = course_match.group(1) if course_match else None
#     course_code = course_match.group(2) if course_match else None
#
#     level_match = re.search(level_pattern, text)
#     skill_level = level_match.group(1) if level_match else None
#
#     return {
#         "Candidate Name": candidate_name,
#         "Guardian Name": guardian_name,
#         "Course Name": course_name,
#         "Course Code": course_code,
#         "Level": skill_level
#     }


def extract_certificate_info(input_path):

    qreader = QReader()
    image = cv2.cvtColor(cv2.imread(input_path), cv2.COLOR_BGR2RGB)
    decoded_texts = qreader.detect_and_decode(image=image)

    string1 = " ".join(decoded_texts)
    print(string1)
    text = string1

    # Define regular expressions to match the required fields
    name_pattern = r"Name:\s*(\w+\s*\w+)"
    father_name_pattern = r"Father Name:\s*(\w+\s*\w+)"
    job_name_pattern = r"Job Name:\s*(.+) \(HSS/(\w+)\)"
    date_pattern = r"Date of Issuance:\s*(\d{2}-[a-zA-Z]{3}-\d{4})"

    # Extract information using regular expressions
    name_match = re.search(name_pattern, text)
    father_name_match = re.search(father_name_pattern, text)
    job_name_match = re.search(job_name_pattern, text)
    date_match = re.search(date_pattern, text)

    # Initialize variables to store the extracted information
    name = father_name = job_name = job_code = date_of_issuance = None

    # Check if matches are found and extract the information
    if name_match:
        name = name_match.group(1)
    if father_name_match:
        father_name = father_name_match.group(1)
    if job_name_match:
        job_name = job_name_match.group(1)
        job_code = job_name_match.group(2)
    if date_match:
        date_of_issuance = date_match.group(1)

    return {
        "QR Name": name,
        "QR Father Name": father_name,
        "QR Job Name": job_name,
        "QR Job Code": job_code,
        "QR Date of Issuance": date_of_issuance
    }

def process_png_files(folder_path, output_file):
    data = []

    # Process each PNG file in the folder
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith('.png'):
            file_path = os.path.join(folder_path, file_name)
            extracted_info = extract_certificate_info(file_path)
            extracted_info['File Name'] = file_name
            data.append(extracted_info)

    # Create a DataFrame and save it to an Excel file
    df = pd.DataFrame(data)

    # Move 'File Name' to the first column
    df = df[['File Name'] + [col for col in df.columns if col != 'File Name']]

    df.to_excel(output_file, index=False)

folder_path = "D:\EY\Cerificate Scanner\Certificate png"
output_file = "D:\EY\Cerificate Scanner\outputqr.xlsx"

process_png_files(folder_path, output_file)
print("Extraction completed and data saved to", output_file)
