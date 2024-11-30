import os
import re
import pandas as pd
from io import BytesIO
from docx import Document
from PyPDF2 import PdfReader
import streamlit as st

# Extract text from a PDF file
def extract_text_from_pdf(pdf_file):
    text = ""
    try:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            text += page.extract_text()
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
    return text

# Extract text from a Word file
def extract_text_from_docx(docx_file):
    text = ""
    try:
        doc = Document(docx_file)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
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

# Streamlit app to upload files and process resumes
def main():
    st.title("Resume Parser")

    # File upload section
    uploaded_files = st.file_uploader("Choose PDF or DOCX files", type=["pdf", "docx"], accept_multiple_files=True)

    if uploaded_files:
        data = []
        
        for uploaded_file in uploaded_files:
            file_type = uploaded_file.type

            # Process the uploaded file based on type
            if file_type == "application/pdf":
                text = extract_text_from_pdf(uploaded_file)
            elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                text = extract_text_from_docx(uploaded_file)
            else:
                st.error(f"Unsupported file format: {file_type}")
                continue

            # Extract relevant information
            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)

        # Create DataFrame and display results
        if data:
            df = pd.DataFrame(data)
            st.write("Parsed Resume Data", df)

            # Allow user to download the parsed data as an Excel file
            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(label="Download Parsed Data", data=output, file_name="parsed_resumes.xlsx", mime="application/vnd.ms-excel")
        else:
            st.warning("No valid data extracted from the resumes.")
    
if __name__ == "__main__":
    main()
