import os
import tempfile
import streamlit as st
from PIL import Image
from PyPDF2 import PdfMerger
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import openpyxl

st.set_page_config(page_title="Universal File Merger", page_icon="📑")
st.title("📑 Universal File Merger")
st.write("Upload PDFs, Excel sheets, PowerPoints, or PNGs to combine them into one final PDF.")

uploaded_files = st.file_uploader(
    "Choose files to merge", 
    type=["pdf", "xlsx", "pptx", "png"], 
    accept_multiple_files=True
)

def excel_to_pdf(file_bytes, temp_pdf_path):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        if any(row):
            data.append([str(cell) if cell is not None else "" for cell in row])
    if not data:
        data = [["Empty Sheet"]]

    doc = SimpleDocTemplate(temp_pdf_path, pagesize=letter)
    t = Table(data)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    doc.build([t])

def image_to_pdf(file_bytes, temp_pdf_path):
    image = Image.open(file_bytes)
    if image.mode != 'RGB':
        image = image.convert('RGB')
    image.save(temp_pdf_path, "PDF", resolution=100.0)

def ppt_to_pdf_text(file_bytes, temp_pdf_path):
    prs = Presentation(file_bytes)
    data = [["Slide #", "Extracted Content"]]
    for idx, slide in enumerate(prs.slides, start=1):
        text_runs = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text_runs.append(paragraph.text)
        slide_text = "\n".join(text_runs).strip()
        data.append([f"Slide {idx}", slide_text if slide_text else "[No Text Content]"])

    doc = SimpleDocTemplate(temp_pdf_path, pagesize=letter)
    t = Table(data, colWidths=[80, 400])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.navy),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
    ]))
    doc.build([t])

if uploaded_files:
    if st.button("Merge Files into PDF", type="primary"):
        merger = PdfMerger()
        temp_dir = tempfile.mkdtemp()
        temp_files = []

        try:
            for i, file in enumerate(uploaded_files):
                ext = os.path.splitext(file.name)[1].lower()
                temp_pdf_path = os.path.join(temp_dir, f"converted_{i}.pdf")

                if ext == ".pdf":
                    merger.append(file)
                elif ext == ".png":
                    image_to_pdf(file, temp_pdf_path)
                    merger.append(temp_pdf_path)
                    temp_files.append(temp_pdf_path)
                elif ext == ".xlsx":
                    excel_to_pdf(file, temp_pdf_path)
                    merger.append(temp_pdf_path)
                    temp_files.append(temp_pdf_path)
                elif ext == ".pptx":
                    ppt_to_pdf_text(file, temp_pdf_path)
                    merger.append(temp_pdf_path)
                    temp_files.append(temp_pdf_path)

            output_pdf_path = os.path.join(temp_dir, "merged_output.pdf")
            merger.write(output_pdf_path)
            merger.close()

            with open(output_pdf_path, "rb") as f:
                st.success("Files merged successfully into PDF!")
                st.download_button(
                    label="Download Master PDF",
                    data=f.read(),
                    file_name="merged_document.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(f"An error occurred: {e}")
