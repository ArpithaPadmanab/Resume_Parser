import io
import os
import re
import pypandoc
from pypandoc import pandoc
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import streamlit as st
from openpyxl import Workbook

# Extract text from PDF file
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text() or ""
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
    return text

# Extract text from DOCX file (including tables)
def extract_text_from_docx(docx_path):
    text = ""
    try:
        doc = Document(docx_path)
        # Extract paragraphs
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"

        # Extract table content
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + " "
            text += "\n"  # Separate rows
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
    return text

# Convert DOC to PDF
def convert_doc_to_pdf(input_path, output_path):
    try:
        pypandoc.convert_file(input_path, "pdf", outputfile=output_path)
        return True
    except Exception as e:
        st.error(f"Error converting DOC to PDF: {e}")
        return False

# Extract text from DOC file
def extract_text_from_doc(doc_path):
    temp_pdf = "temp_output.pdf"
    if convert_doc_to_pdf(doc_path, temp_pdf):
        text = extract_text_from_pdf(temp_pdf)
        os.remove(temp_pdf)  # Clean up
        return text
    return ""

# Extract information using regex
def extract_info(text):
    info = {
        "Name": None,
        "Email": None,
        "Phone": None,
        "Education": None,
        "Skills": None,
        "Experience": None,
        "Position": None,  # New column for position
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

    # Extract education
    education_pattern = r"(B\.E|B\.Tech|B\.Sc|B\.Com|M\.Tech|M\.Sc|PhD|MBA|Bachelor|Master|Diploma)"
    education_match = re.search(education_pattern, text, re.IGNORECASE)
    if education_match:
        info["Education"] = education_match.group(0)

    # Extract skills
    skills_keywords = [
        "C++", "C", ".NET", "Python", "Java", "SQL", "Machine Learning", "Data Science",
        "Tableau", "PowerBI", "PLC", "DCS", "SCADA", "AutoCAD", "P2P", "O2C", "SCM", "MM",
        "SAP", "Robo", "BiW", "SolidWorks", "Mechanical Design", "Electrical Design", "E Plan",
        "LV", "MV", "LT", "MT", "EBASE", "800xA"
    ]
    skills_found = [skill for skill in skills_keywords if skill.lower() in text.lower()]
    info["Skills"] = ", ".join(skills_found)

    # Extract experience
    experience_pattern = r"(\d+\s+(years?|months?)\s+experience)"
    experience_match = re.search(experience_pattern, text, re.IGNORECASE)
    if experience_match:
        info["Experience"] = experience_match.group(0)

    # Determine position based on keywords
    position_keywords = {
        "Software Engineer": ["Python", "Java", "C++", ".NET"],
        "Data Scientist": ["Machine Learning", "Data Science", "SQL", "Tableau", "PowerBI"],
        "Mechanical Engineer": ["AutoCAD", "SolidWorks", "Mechanical Design"],
        "Electrical Engineer": ["Electrical Design", "E Plan", "LV", "MV", "800xA"],
        "SAP Consultant": ["SAP", "P2P", "O2C", "SCM", "MM"],
    }

    for position, keywords in position_keywords.items():
        if any(keyword.lower() in text.lower() for keyword in keywords):
            info["Position"] = position
            break

    return info

# Streamlit App
st.set_page_config(page_title="Resume Tracker", layout="wide")

col1, col2 = st.columns([1, 2])  # Adjust the width ratio if needed

# Add an image in the first column
with col1:
    try:
        st.image("logo.jpeg", use_column_width=True)
    except FileNotFoundError:
        st.write("Logo not found. Please ensure 'logo.jpeg' is in the app directory.")

# Add text in the second column
with col2:
    st.title("Resume Tracker")

uploaded_files = st.file_uploader("Upload resumes", type=["pdf", "docx", "doc"], accept_multiple_files=True)

if uploaded_files:
    data = []
    progress = st.progress(0)
    for idx, uploaded_file in enumerate(uploaded_files):
        text = ""

        # Handle different file types
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith(".docx"):
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            text = extract_text_from_docx(temp_path)
            os.remove(temp_path)
        elif uploaded_file.name.endswith(".doc"):
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            text = extract_text_from_doc(temp_path)
            os.remove(temp_path)

        if text:
            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)
        else:
            st.warning(f"Could not process file: {uploaded_file.name}")

        progress.progress((idx + 1) / len(uploaded_files))

    df = pd.DataFrame(data)
    st.dataframe(df)

    def convert_df_to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Resumes")
        return output.getvalue()

    if not df.empty:
        excel_data = convert_df_to_excel(df)
        st.download_button(
            "Download Excel",
            data=excel_data,
            file_name="Resume_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
