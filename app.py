import os
import io
import re
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import streamlit as st
from openpyxl import Workbook
from transformers import pipeline
import torch

# Force transformers to use CPU only
device = -1  # -1 means CPU

# Load NLP models
ner_pipeline = pipeline("ner", model="dbmdz/bert-large-cased-finetuned-conll03-english", device=device)
skill_extractor = pipeline("zero-shot-classification", model="facebook/bart-large-mnli", device=device)

# Utility Functions
def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file."""
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text() or ""
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
    return text

def extract_text_from_docx(docx_path):
    """Extract text from a DOCX file."""
    text = ""
    try:
        doc = Document(docx_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
    return text

def extract_info(text):
    """Extract relevant information using NLP models."""
    info = {
        "Name": None,
        "Email": None,
        "Phone": None,
        "Education": None,
        "Skills": None,
        "Experience": None,
        "Position": None,
    }

    # Extract name (Using Named Entity Recognition)
    entities = ner_pipeline(text)
    for ent in entities:
        if ent["entity"] == "B-PER":
            info["Name"] = ent["word"]
            break

    # Extract email
    email_pattern = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
    email_match = re.search(email_pattern, text)
    if email_match:
        info["Email"] = email_match.group(0)

    # Extract phone
    phone_pattern = r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b'
    phone_match = re.search(phone_pattern, text)
    if phone_match:
        info["Phone"] = phone_match.group(0)

    # Extract education
    education_pattern = r"(B\.E|B\.Tech|B\.Sc|B\.Com|M\.Tech|M\.Sc|PhD|MBA|Bachelor|Master|Diploma)"
    education_match = re.search(education_pattern, text, re.IGNORECASE)
    if education_match:
        info["Education"] = education_match.group(0)
    
    # Extract skills using NLP Model
    skills_keywords = [
        "C++", "C", ".NET", "Python", "Java", "SQL", "Machine Learning", "Data Science",
        "Tableau", "PowerBI", "PLC", "DCS", "SCADA", "AutoCAD", "P2P", "O2C", "SCM", "MM",
        "SAP", "Robo", "BiW", "SolidWorks", "Mechanical Design", "Electrical Design", "E Plan",
        "LV", "MV", "LT", "MT", "EBASE", "800xA", "B.Com"
    ]
    skills = skill_extractor(text, skills_keywords, multi_label=True)
    info["Skills"] = ", ".join([label for label, score in zip(skills["labels"], skills["scores"]) if score > 0.5])
    
    # Extract experience
    experience_pattern = r"(\d+\s+(years?|months?)\s+experience)"
    experience_match = re.search(experience_pattern, text, re.IGNORECASE)
    if experience_match:
        info["Experience"] = experience_match.group(0)
    
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
        if any(keyword.lower() in info["Skills"].lower() for keyword in keywords):
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
col1, col2 = st.columns([1, 2])  # Adjust the width ratio if needed

# Add an image in the first column
with col1:
    st.image(
        "logo.jpeg"
    )
# Add text in the second column
with col2:
   st.title("RESUME TRACKER")

# File Uploader
uploaded_files = st.file_uploader("Upload resumes", type=["pdf", "docx"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        text = ""
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith(".docx"):
            with open(f"temp_{uploaded_file.name}", "wb") as f:
                f.write(uploaded_file.getbuffer())
            text = extract_text_from_docx(f"temp_{uploaded_file.name}")
            os.remove(f"temp_{uploaded_file.name}")
        
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
