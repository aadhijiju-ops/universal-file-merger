import io
import streamlit as st
from PIL import Image
from PyPDF2 import PdfMerger
from pptx import Presentation
import openpyxl

st.set_page_config(page_title="Multi-Format File Merger", page_icon="🧰")
st.title("🧰 Multi-Format File Merger")
st.write("Select a file type below to merge your files into their original format.")

# Create 4 distinct tabs for each file type
tab_pdf, tab_png, tab_ppt, tab_excel = st.tabs(["📄 PDF Merger", "🖼️ PNG Merger", "📊 PPT Merger", "📊 Excel Merger"])

# ---------------------------------------------------------
# 1. PDF MERGER
# ---------------------------------------------------------
with tab_pdf:
    st.subheader("Merge PDFs into a single PDF")
    pdf_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True, key="pdf_uploader")
    
    if pdf_files and len(pdf_files) >= 2:
        if st.button("Merge PDFs", type="primary", key="btn_pdf"):
            try:
                merger = PdfMerger()
                for pdf in pdf_files:
                    merger.append(pdf)
                
                output = io.BytesIO()
                merger.write(output)
                merger.close()
                output.seek(0)
                
                st.success("PDFs merged successfully!")
                st.download_button(
                    label="Download Merged PDF",
                    data=output,
                    file_name="merged_document.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Error merging PDFs: {e}")
    elif pdf_files:
        st.info("Please upload at least 2 PDF files to merge.")

# ---------------------------------------------------------
# 2. PNG MERGER
# ---------------------------------------------------------
with tab_png:
    st.subheader("Stitch PNGs into a single long PNG")
    png_files = st.file_uploader("Upload PNG images", type=["png"], accept_multiple_files=True, key="png_uploader")
    
    if png_files and len(png_files) >= 2:
        if st.button("Merge PNGs", type="primary", key="btn_png"):
            try:
                images = [Image.open(img) for img in png_files]
                
                # Combine images vertically
                max_width = max(img.width for img in images)
                total_height = sum(img.height for img in images)
                
                merged_image = Image.new("RGBA", (max_width, total_height))
                
                y_offset = 0
                for img in images:
                    merged_image.paste(img, (0, y_offset))
                    y_offset += img.height
                
                output = io.BytesIO()
                merged_image.save(output, format="PNG")
                output.seek(0)
                
                st.success("PNGs merged successfully!")
                st.download_button(
                    label="Download Merged PNG",
                    data=output,
                    file_name="merged_image.png",
                    mime="image/png"
                )
            except Exception as e:
                st.error(f"Error merging PNGs: {e}")
    elif png_files:
        st.info("Please upload at least 2 PNG files to merge.")

# ---------------------------------------------------------
# 3. PPT MERGER
# ---------------------------------------------------------
with tab_ppt:
    st.subheader("Merge PowerPoint slides into a single PPTX")
    ppt_files = st.file_uploader("Upload PPTX files", type=["pptx"], accept_multiple_files=True, key="ppt_uploader")
    
    if ppt_files and len(ppt_files) >= 2:
        if st.button("Merge PPTs", type="primary", key="btn_ppt"):
            try:
                base_prs = Presentation(ppt_files[0])
                
                for ppt in ppt_files[1:]:
                    sub_prs = Presentation(ppt)
                    for slide in sub_prs.slides:
                        blank_layout = base_prs.slide_layouts[6]
                        new_slide = base_prs.slides.add_slide(blank_layout)
                        for shape in slide.shapes:
                            el = shape.element
                            new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')
                
                output = io.BytesIO()
                base_prs.save(output)
                output.seek(0)
                
                st.success("PowerPoint files merged successfully!")
                st.download_button(
                    label="Download Merged PPTX",
                    data=output,
                    file_name="merged_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"Error merging PPTs: {e}")
    elif ppt_files:
        st.info("Please upload at least 2 PPTX files to merge.")

# ---------------------------------------------------------
# 4. EXCEL MERGER
# ---------------------------------------------------------
with tab_excel:
    st.subheader("Combine Excel sheets into a single Excel workbook")
    excel_files = st.file_uploader("Upload Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="excel_uploader")
    
    if excel_files and len(excel_files) >= 2:
        if st.button("Merge Excel Files", type="primary", key="btn_excel"):
            try:
                merged_wb = openpyxl.Workbook()
                # Remove default sheet
                merged_wb.remove(merged_wb.active)
                
                for idx, fx in enumerate(excel_files, start=1):
                    wb = openpyxl.load_workbook(fx, data_only=True)
                    for sheet_name in wb.sheetnames:
                        source_sheet = wb[sheet_name]
                        new_sheet_title = f"File{idx}_{sheet_name}"
                        target_sheet = merged_wb.create_sheet(title=new_sheet_title[:31]) # Max sheet name length is 31
                        
                        for row in source_sheet.iter_rows(values_only=True):
                            target_sheet.append(row)
                
                output = io.BytesIO()
                merged_wb.save(output)
                output.seek(0)
                
                st.success("Excel files merged successfully!")
                st.download_button(
                    label="Download Merged Excel",
                    data=output,
                    file_name="merged_sheets.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error merging Excel files: {e}")
    elif excel_files:
        st.info("Please upload at least 2 Excel files to merge.")
