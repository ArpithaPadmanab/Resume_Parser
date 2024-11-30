import os
import re
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader

# Extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text()
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
    return text

# Extract text from a Word file
def extract_text_from_docx(docx_path):
    text = ""
    try:
        doc = Document(docx_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        print(f"Error reading DOCX {docx_path}: {e}")
    return text

# Extract relevant information using regex
def extract_info(text):
    info = {
        "Name": None,
        "Email": None,
        "Phone": None,
        "Education": None,
        "Skills": None,
        "Experience": None,
    }

    # Extract email
    email_pattern = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
    email_match = re.search(email_pattern, text)
    if email_match:
        info["Email"] = email_match.group(0)

    # Extract phone number
    phone_pattern = r'\b\d{10}\b'
    phone_match = re.search(phone_pattern, text)
    if phone_match:
        info["Phone"] = phone_match.group(0)

    # Extract education (sample pattern)
    education_pattern = r"(B\.Tech|B\.Sc|M\.Tech|M\.Sc|PhD|MBA)"
    education_match = re.search(education_pattern, text, re.IGNORECASE)
    if education_match:
        info["Education"] = education_match.group(0)

    # Extract skills (simple keyword matching)
    skills_keywords = ["Python", "Java", "SQL", "Machine Learning", "Data Science"]
    skills_found = [skill for skill in skills_keywords if skill.lower() in text.lower()]
    info["Skills"] = ", ".join(skills_found)

    # Extract experience (sample pattern)
    experience_pattern = r"(\d+ years? experience)"
    experience_match = re.search(experience_pattern, text, re.IGNORECASE)
    if experience_match:
        info["Experience"] = experience_match.group(0)

    return info

def process_resumes(input_folder, output_file):
    data = []

    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder, filename)
        if filename.endswith(".pdf"):
            text = extract_text_from_pdf(file_path)
        elif filename.endswith(".docx"):
            text = extract_text_from_docx(file_path)
        else:
            print(f"Unsupported file format: {filename}")
            continue

        # Extract information
        info = extract_info(text)
        info["Filename"] = filename
        data.append(info)

    # Create a DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")

if __name__ == "__main__":
    input_folder = "C:\\Users\\india\\OneDrive\\Desktop\\Resume parasing\\Resume"  # Update with your folder path
    output_file = "C:\\Users\\india\\OneDrive\\Desktop\\Resume parasing\\output.xlsx"            # Update with desired output file name
    
    process_resumes(input_folder, output_file)
