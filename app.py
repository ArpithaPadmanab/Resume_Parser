import io
import os
import re
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import streamlit as st

# Google Drive authentication
def authenticate_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()  # Creates a local webserver and auto-handles authentication
    return GoogleDrive(gauth)

drive = authenticate_drive()

# Upload file to Google Drive
def upload_to_drive(file_content, file_name):
    try:
        file_drive = drive.CreateFile({'title': file_name})
        file_drive.SetContentString(file_content)
        file_drive.Upload()
        return file_drive['id']
    except Exception as e:
        st.error(f"Error uploading file to Google Drive: {e}")
        return None

# Generate Google Drive file link
def generate_drive_link(file_id):
    return f"https://drive.google.com/file/d/{file_id}/view"

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
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + " "
            text += "\n"
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
    return text

# Extract information using regex
def extract_info(text):
    info = {
        "Name": None,
        "Email": None,
        "Phone": None,
        "Education": None,
        "Skills": None,
        "Experience": None,
    }

    email_pattern = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
    email_match = re.search(email_pattern, text)
    if email_match:
        info["Email"] = email_match.group(0)

    phone_pattern = r'\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b'
    phone_match = re.search(phone_pattern, text)
    if phone_match:
        info["Phone"] = phone_match.group(0)

    lines = text.split("\n")
    if lines:
        info["Name"] = lines[0].strip()

    education_pattern = r"(B\.Tech|B\.Sc|M\.Tech|M\.Sc|PhD|MBA)"
    education_match = re.search(education_pattern, text, re.IGNORECASE)
    if education_match:
        info["Education"] = education_match.group(0)

    skills_keywords = ["Python", "Java", "SQL", "Machine Learning", "Data Science"]
    skills_found = [skill for skill in skills_keywords if skill.lower() in text.lower()]
    info["Skills"] = ", ".join(skills_found)

    experience_pattern = r"(\d+ years? experience)"
    experience_match = re.search(experience_pattern, text, re.IGNORECASE)
    if experience_match:
        info["Experience"] = experience_match.group(0)

    return info

# Streamlit App
col1, col2 = st.columns([1, 2])

with col1:
    st.image("logo.jpeg")

with col2:
    st.title("RESUME TRACKER")

uploaded_files = st.file_uploader("Upload resumes", type=["pdf", "docx"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        text = ""
        file_content = uploaded_file.read().decode("latin1")
        drive_id = upload_to_drive(file_content, uploaded_file.name)

        if drive_id:
            link = generate_drive_link(drive_id)

            if uploaded_file.name.endswith(".pdf"):
                temp_path = f"temp_{uploaded_file.name}"
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                text = extract_text_from_pdf(temp_path)
                os.remove(temp_path)
            elif uploaded_file.name.endswith(".docx"):
                temp_path = f"temp_{uploaded_file.name}"
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                text = extract_text_from_docx(temp_path)
                os.remove(temp_path)

            if text:
                info = extract_info(text)
                info["Filename"] = link
                data.append(info)
            else:
                st.warning(f"Could not process file: {uploaded_file.name}")

    df = pd.DataFrame(data)
    st.dataframe(df)

    def convert_df(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Resumes")
        return output.getvalue()

    if not df.empty:
        excel_data = convert_df(df)
        st.download_button(
            "Download Excel",
            data=excel_data,
            file_name="Resume_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
