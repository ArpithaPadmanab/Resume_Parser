import os
import io
import re
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import streamlit as st
from openpyxl import Workbook
from transformers import pipeline

# ---------------------------------------------------------
# ✅ Caching models for performance
# ---------------------------------------------------------

@st.cache_resource
def load_ner_pipeline():
    return pipeline("ner", model="dbmdz/bert-large-cased-finetuned-conll03-english", device=-1)
    # ➕ Replace with "dslim/bert-base-NER" for faster loads if needed

@st.cache_resource
def load_skill_extractor():
    return pipeline("zero-shot-classification", model="facebook/bart-large-mnli", device=-1)
    # ➕ Replace with "joeddav/xlm-roberta-large-xnli" for multilingual + faster inference

# Load models once with a spinner
with st.spinner("🔁 Loading NLP models..."):
    ner_pipeline = load_ner_pipeline()
    skill_extractor = load_skill_extractor()

# ---------------------------------------------------------
# 📄 File Extractors
# ---------------------------------------------------------

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text() or ""
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
    return text

def extract_text_from_docx(docx_path):
    text = ""
    try:
        doc = Document(docx_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
    return text

# ---------------------------------------------------------
# 📊 Info Extractor
# ---------------------------------------------------------

def extract_info(text):
    info = {
        "Name": None,
        "Email": None,
        "Phone": None,
        "Education": None,
        "Skills": None,
        "Experience": None,
        "Position": None,
    }

    # 🔎 Name extraction
    entities = ner_pipeline(text)
    for ent in entities:
        if ent["entity"] == "B-PER":
            info["Name"] = ent["word"]
            break

    # 📧 Email
    email_match = re.search(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', text)
    if email_match:
        info["Email"] = email_match.group(0)

    # 📞 Phone
    phone_match = re.search(r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b', text)
    if phone_match:
        info["Phone"] = phone_match.group(0)

    # 🎓 Education
    education_match = re.search(r"(B\.E|B\.Tech|B\.Sc|B\.Com|M\.Tech|M\.Sc|PhD|MBA|Bachelor|Master|Diploma)", text, re.IGNORECASE)
    if education_match:
        info["Education"] = education_match.group(0)

    # 🛠️ Skills
    skills_keywords = [
        "C++", "C", ".NET", "Python", "Java", "SQL", "Machine Learning", "Data Science",
        "Tableau", "PowerBI", "PLC", "DCS", "SCADA", "AutoCAD", "P2P", "O2C", "SCM", "MM",
        "SAP", "Robo", "BiW", "SolidWorks", "Mechanical Design", "Electrical Design", "E Plan",
        "LV", "MV", "LT", "MT", "EBASE", "800xA", "B.Com"
    ]
    skills = skill_extractor(text, skills_keywords, multi_label=True)
    info["Skills"] = ", ".join([label for label, score in zip(skills["labels"], skills["scores"]) if score > 0.5])

    # 🧑‍💼 Experience
    experience_match = re.search(r"(\d+\s+(years?|months?)\s+experience)", text, re.IGNORECASE)
    if experience_match:
        info["Experience"] = experience_match.group(0)

    # 📌 Position Mapping
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

# ---------------------------------------------------------
# 📤 Excel Export
# ---------------------------------------------------------

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resumes")
    return output.getvalue()

# ---------------------------------------------------------
# 🎯 Streamlit UI
# ---------------------------------------------------------

st.set_page_config(page_title="Resume Tracker", layout="wide")

# Header
col1, col2 = st.columns([1, 2])
with col1:
    st.image("logo.jpeg", use_column_width=True)
with col2:
    st.title("📂 Resume Tracker")

# Upload Files
uploaded_files = st.file_uploader("Upload resumes", type=["pdf", "docx"], accept_multiple_files=True)

# Main Logic
if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        text = ""
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith(".docx"):
            temp_path = f"temp_{uploaded_file.name}"
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            text = extract_text_from_docx(temp_path)
            os.remove(temp_path)

        if text:
            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)
        else:
            st.warning(f"⚠️ Could not process file: {uploaded_file.name}")

    # Display results
    df = pd.DataFrame(data)
    st.dataframe(df)

    # Download button
    if not df.empty:
        excel_data = convert_df_to_excel(df)
        st.download_button(
            "📥 Download Excel",
            data=excel_data,
            file_name="Resume_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
