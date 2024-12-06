import io
import os
import re
import pypandoc
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import streamlit as st

# Helper Functions
def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file."""
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text() or ""
    except Exception as e:
        st.error(f"Error reading PDF {pdf_path}: {e}")
    return text


def extract_text_from_docx(docx_path):
    """Extract text from DOCX file."""
    text = ""
    try:
        doc = Document(docx_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        st.error(f"Error reading DOCX {docx_path}: {e}")
    return text


def convert_doc_to_pdf(input_path, output_path):
    """Convert DOC file to PDF."""
    try:
        pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)
        return True
    except Exception as e:
        st.error(f"Error converting DOC to PDF: {e}")
        return False


def extract_text_from_doc(doc_path):
    """Extract text from DOC file by converting to PDF."""
    temp_pdf = "temp_output.pdf"
    if convert_doc_to_pdf(doc_path, temp_pdf):
        text = extract_text_from_pdf(temp_pdf)
        os.remove(temp_pdf)  # Clean up temporary file
        return text
    return ""


def extract_info(text):
    """Extract relevant information from the text."""
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
    phone_pattern = r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b'
    phone_match = re.search(phone_pattern, text)
    if phone_match:
        info["Phone"] = phone_match.group(0)

    # Extract name (first line heuristic)
    lines = text.split("\n")
    if lines:
        info["Name"] = lines[0].strip()

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


# Streamlit App
st.title("Resume Parsing Application")

uploaded_files = st.file_uploader("Upload resumes", type=["pdf", "docx", "doc"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        text = ""

        # Handle different file types
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith(".docx"):
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            text = extract_text_from_docx(temp_path)
            os.remove(temp_path)  # Clean up temporary file
        elif uploaded_file.name.endswith(".doc"):
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            text = extract_text_from_doc(temp_path)
            os.remove(temp_path)  # Clean up temporary file

        # Extract info and append to data
        if text:
            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)
        else:
            st.warning(f"Could not process file: {uploaded_file.name}")

    # Display results in a DataFrame
    df = pd.DataFrame(data)
    st.dataframe(df)

    # Provide download option
    def convert_df(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Resumes")
        return output.getvalue()

    if not df.empty:
        excel_data = convert_df(df)
        st.download_button("Download Excel", data=excel_data, file_name="Resume_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
