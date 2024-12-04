import io
import os
import re
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import streamlit as st

# Install python-docx and comtypes for handling docx and doc files.
try:
    import comtypes.client
except ImportError:
    comtypes = None

# Helper Functions

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

# Extract text from a DOCX file
def extract_text_from_docx(docx_file):
    text = ""
    try:
        if isinstance(docx_file, (str, bytes)):  # File path or bytes
            doc = Document(docx_file)
        else:  # File-like object (e.g., from st.file_uploader)
            doc = Document(docx_file)

        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        print(f"Error reading DOCX file: {e}")
    return text

# Extract text from a DOC file (requires comtypes)
def extract_text_from_doc(doc_file):
    text = ""
    if not comtypes:
        print("comtypes is required for reading .doc files.")
        return text

    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_file.name)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
    except Exception as e:
        print(f"Error reading DOC file: {e}")
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

    # Extract education
    education_pattern = r"(B\.Tech|B\.Sc|M\.Tech|M\.Sc|PhD|MBA)"
    education_match = re.search(education_pattern, text, re.IGNORECASE)
    if education_match:
        info["Education"] = education_match.group(0)

    # Extract skills
    skills_keywords = ["Python", "Java", "SQL", "Machine Learning", "Data Science"]
    skills_found = [skill for skill in skills_keywords if skill.lower() in text.lower()]
    info["Skills"] = ", ".join(skills_found)

    # Extract experience
    experience_pattern = r"(\d+ years? experience)"
    experience_match = re.search(experience_pattern, text, re.IGNORECASE)
    if experience_match:
        info["Experience"] = experience_match.group(0)

    return info

# Streamlit App
st.title("Resume Parsing Application")

uploaded_files = st.file_uploader("Upload resumes", type=["pdf", "docx", "doc"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        text = ""
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith(".docx"):
            text = extract_text_from_docx(uploaded_file)
        elif uploaded_file.name.endswith(".doc"):
            text = extract_text_from_doc(uploaded_file)
        else:
            st.warning(f"Unsupported file type: {uploaded_file.name}")

        if text:
            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)

    # Display extracted information
    if data:
        df = pd.DataFrame(data)
        st.dataframe(df)

        # Download the DataFrame as an Excel file
        @st.cache_data
        def convert_df(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            processed_data = output.getvalue()
            return processed_data

        excel_data = convert_df(df)
        st.download_button("Download Excel", data=excel_data, file_name="Resume_Data.xlsx")
