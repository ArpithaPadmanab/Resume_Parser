import io
import os
import re
import tempfile
from docx import Document
from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import streamlit as st
import pandas as pd
import pythoncom
from win32com import client

# Convert DOC to PDF
def convert_doc_to_pdf(doc_path, output_path):
    try:
        pythoncom.CoInitialize()  # Ensure COM library is initialized
        word = client.Dispatch("Word.Application")
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(output_path, FileFormat=17)  # 17 corresponds to wdFormatPDF
        doc.Close()
        word.Quit()
        return output_path
    except Exception as e:
        raise Exception(f"Error converting DOC to PDF: {e}")

# Convert DOCX to PDF with Unicode support
def convert_docx_to_pdf(docx_path, output_path):
    try:
        doc = Document(docx_path)
        pdf_canvas = canvas.Canvas(output_path, pagesize=letter)
        width, height = letter
        margin = 72
        y_position = height - margin

        pdf_canvas.setFont("Helvetica", 12)
        for paragraph in doc.paragraphs:
            text = paragraph.text
            if text.strip():
                lines = pdf_canvas.beginText(margin, y_position)
                lines.setTextOrigin(margin, y_position)
                lines.textLines(text)
                pdf_canvas.drawText(lines)
                y_position -= 20
                if y_position < margin:
                    pdf_canvas.showPage()
                    y_position = height - margin

        pdf_canvas.save()
        return output_path
    except Exception as e:
        raise Exception(f"Error converting DOCX to PDF: {e}")

# Extract text from PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text += page.extract_text()
    except Exception as e:
        raise Exception(f"Error reading PDF: {e}")
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

    # Extract name (assumes name is the first line or near the top)
    lines = text.split("\n")
    for line in lines:
        if len(line.split()) > 1:  # Likely a name
            info["Name"] = line.strip()
            break

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

# Streamlit app
st.title("Resume Parsing Application")

uploaded_files = st.file_uploader("Upload resumes (PDF, DOCX, DOC)", type=["pdf", "docx", "doc"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as temp_file:
            temp_file.write(uploaded_file.read())
            temp_file_path = temp_file.name

        try:
            if uploaded_file.name.endswith(".pdf"):
                text = extract_text_from_pdf(temp_file_path)
            elif uploaded_file.name.endswith(".docx"):
                pdf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                convert_docx_to_pdf(temp_file_path, pdf_path)
                text = extract_text_from_pdf(pdf_path)
            elif uploaded_file.name.endswith(".doc"):
                pdf_path = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                convert_doc_to_pdf(temp_file_path, pdf_path)
                text = extract_text_from_pdf(pdf_path)

            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")

    df = pd.DataFrame(data)
    st.dataframe(df)

    # Download the DataFrame as an Excel file
    @st.cache_data
    def convert_df_to_excel(dataframe):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False, sheet_name='Resumes')
        return output.getvalue()

    if not df.empty:
        excel_data = convert_df_to_excel(df)
        st.download_button("Download Excel", data=excel_data, file_name="Resume_Data.xlsx")
