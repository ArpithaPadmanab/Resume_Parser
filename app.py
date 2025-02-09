import os
import io
import re
import spacy
import pandas as pd
import streamlit as st
from docx import Document
from PyPDF2 import PdfReader
from openpyxl import Workbook

# Load spaCy NLP model for better name extraction
nlp = spacy.load("en_core_web_sm")

# Utility Functions
def extract_text_from_pdf(pdf_file):
    """Extract text from a PDF file."""
    text = ""
    try:
        reader = PdfReader(pdf_file)
        for page in reader.pages:
            text += page.extract_text() or ""
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
    return text.strip()

def extract_text_from_docx(docx_file):
    """Extract text from a DOCX file."""
    text = ""
    try:
        doc = Document(io.BytesIO(docx_file.getvalue()))
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + " "
            text += "\n"
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
    return text.strip()

def extract_name(text):
    """Extract name using NLP."""
    doc = nlp(text)
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            return ent.text
    return None

def extract_info(text):
    """Extract relevant information from resume text."""
    info = {
        "Name": extract_name(text),
        "Email": None,
        "Phone": None,
        "Education": None,
        "Skills": None,
        "Experience": None,
        "Position": None,
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

    # Extract education (multiple matches)
    education_pattern = r"(B\.E|B\.Tech|B\.Sc|B\.Com|M\.Tech|M\.Sc|PhD|MBA|Bachelor|Master|Diploma)"
    education_matches = re.findall(education_pattern, text, re.IGNORECASE)
    if education_matches:
        info["Education"] = ", ".join(set(education_matches))

    # Extract skills
    skills_keywords = [
        "C++", "C", ".NET", "Python", "Java", "SQL", "Machine Learning", "Data Science",
        "Tableau", "PowerBI", "PLC", "DCS", "SCADA", "AutoCAD", "P2P", "O2C", "SCM", "MM",
        "SAP", "Robo", "BiW", "SolidWorks", "Mechanical Design", "Electrical Design", "E Plan",
        "LV", "MV", "LT", "MT", "EBASE", "800xA", "B.Com"
    ]
    skills_found = [skill for skill in skills_keywords if skill.lower() in text.lower()]
    info["Skills"] = ", ".join(skills_found) if skills_found else None

    # Extract experience (multiple formats)
    experience_pattern = r"(\d+)\s+(years?|months?)\s+experience"
    experience_matches = re.findall(experience_pattern, text, re.IGNORECASE)
    if experience_matches:
        info["Experience"] = ", ".join([f"{match[0]} {match[1]}" for match in experience_matches])

    # Assign position
    position_keywords = {
        "Finance": ["B.Com"],
        "Purchase Engineer": ["SCM", "SAP", "MM"],
        "Order Management Associate": ["B.Com", "SAP"],
        "Data Analytics": ["Machine Learning", "Data Science", "SQL", "Tableau", "PowerBI"],
        "Software Engineer": ["Python", "Java", "C++", ".NET"],
        "800xA": ["800xA"],
        "SAP Consultant": ["SAP", "LV", "MV", "LT", "MT"],
        "Electrical Engineer": ["Electrical Design", "E Plan", "EBASE"],
        "Mechanical Design": ["SolidWorks", "Mechanical Design"],
        "Automation Engineer": ["PLC", "DCS", "SCADA"],
        "AutoCAD": ["AutoCAD"],
        "Sales Support Engineer": ["P2P", "O2C"],
        "Robotics Programmer": ["Robo"],
        "BiW": ["BiW"],
    }

    for position, keywords in position_keywords.items():
        if any(keyword.lower() in text.lower() for keyword in keywords):
            info["Position"] = position
            break

    return info

def convert_df_to_excel(df):
    """Convert a DataFrame to an Excel file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resumes")
    return output.getvalue()

# Streamlit App
st.set_page_config(page_title="Resume Tracker", layout="wide")

# UI Components
col1, col2 = st.columns([1, 2])

# Add image with corrected parameter
with col1:
    st.image("logo.jpeg", caption="Company Logo", use_container_width=True)

# Add title
with col2:
    st.title("Resume Tracker")

# File Uploader
uploaded_files = st.file_uploader("Upload resumes", type=["pdf", "docx"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        text = ""
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith(".docx"):
            text = extract_text_from_docx(uploaded_file)

        if text:
            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)
        else:
            st.warning(f"Could not process file: {uploaded_file.name}")

    # Display DataFrame
    df = pd.DataFrame(data)
    st.dataframe(df)

    # Download Button
    if not df.empty:
        excel_data = convert_df_to_excel(df)
        st.download_button(
            "Download Excel",
            data=excel_data,
            file_name="Resume_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
