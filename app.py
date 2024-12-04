import os
import re
import io
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import streamlit as st

try:
    import comtypes.client
except ImportError:
    comtypes = None

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

    # Extract skills (simple keyword matching)
    skills_keywords = ["Python", "Java", "SQL", "Machine Learning", "Data Science"]
    skills_found = [skill for skill in skills_keywords if skill.lower() in text.lower()]
    info["Skills"] = ", ".join(skills_found)

    # Extract experience
    experience_pattern = r"(\d+ years? experience)"
    experience_match = re.search(experience_pattern, text, re.IGNORECASE)
    if experience_match:
        info["Experience"] = experience_match.group(0)

    return info

# Convert DOCX to PDF
def convert_docx_to_pdf(docx_path, output_path):
    from fpdf import FPDF

    try:
        doc = Document(docx_path)
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        for paragraph in doc.paragraphs:
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, paragraph.text)

        pdf.output(output_path)
        return output_path
    except Exception as e:
        st.error(f"Error converting DOCX to PDF: {e}")
        return None

# Convert DOC to PDF
def convert_doc_to_pdf(doc_path, output_path):
    if comtypes is None:
        st.error("comtypes module is required for processing .doc files but is not installed.")
        return None

    try:
        word = comtypes.client.CreateObject("Word.Application")
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF format
        doc.Close()
        word.Quit()
        return output_path
    except Exception as e:
        st.error(f"Error converting DOC to PDF: {e}")
        return None

# Extract text from PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            if page.extract_text():
                text += page.extract_text()
    except Exception as e:
        st.error(f"Error reading PDF {pdf_path}: {e}")
    return text

# Streamlit UI
st.title("Resume Parsing with DOC/DOCX to PDF Conversion")

uploaded_files = st.file_uploader("Upload resumes (.pdf, .docx, .doc)", type=["pdf", "docx", "doc"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for uploaded_file in uploaded_files:
        text = ""
        temp_dir = os.getcwd()
        temp_file_path = os.path.join(temp_dir, uploaded_file.name)

        # Save uploaded file temporarily
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Handle file type
        if uploaded_file.name.endswith(".pdf"):
            text = extract_text_from_pdf(temp_file_path)
        elif uploaded_file.name.endswith(".docx"):
            temp_pdf_path = temp_file_path.replace(".docx", ".pdf")
            pdf_path = convert_docx_to_pdf(temp_file_path, temp_pdf_path)
            if pdf_path:
                text = extract_text_from_pdf(pdf_path)
        elif uploaded_file.name.endswith(".doc"):
            temp_pdf_path = temp_file_path.replace(".doc", ".pdf")
            pdf_path = convert_doc_to_pdf(temp_file_path, temp_pdf_path)
            if pdf_path:
                text = extract_text_from_pdf(pdf_path)

        # Extract information
        if text:
            info = extract_info(text)
            info["Filename"] = uploaded_file.name
            data.append(info)

    # Display extracted information
    df = pd.DataFrame(data)
    st.dataframe(df)

    # Download as Excel
    def convert_df_to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Resumes')
        return output.getvalue()

    if not df.empty:
        excel_data = convert_df_to_excel(df)
        st.download_button("Download Excel", data=excel_data, file_name="Parsed_Resumes.xlsx")
