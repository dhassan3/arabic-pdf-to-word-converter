import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import zipfile
from pathlib import Path
import arabic_reshaper
from bidi.algorithm import get_display
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

st.set_page_config(page_title="Arabic PDF → Word", page_icon="📄")
st.title("🇸🇦 Arabic PDF to Word Converter")
st.markdown("""
Upload your Arabic PDFs and get perfectly formatted Word files.  
Supports RTL text, connected letters, tables, and images.
""")

# Sidebar options
st.sidebar.header("Conversion Options")
convert_all = st.sidebar.checkbox("Convert all pages", value=True)
if not convert_all:
    col1, col2 = st.sidebar.columns(2)
    start_page = col1.number_input("Start page", min_value=1, value=1, step=1)
    end_page = col2.number_input("End page", min_value=1, value=20, step=1)
    if start_page > end_page:
        st.sidebar.error("Start page must be ≤ End page")
        st.stop()
else:
    start_page = end_page = None

# File uploader
uploaded_files = st.file_uploader(
    "Choose PDF file(s)",
    type="pdf",
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("👆 Upload one or more Arabic PDFs to get started.")
    st.stop()

with tempfile.TemporaryDirectory() as temp_dir:
    temp_path = Path(temp_dir)
    output_files = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uploaded_file in enumerate(uploaded_files):
        filename_base = Path(uploaded_file.name).stem
        input_pdf = temp_path / uploaded_file.name
        with open(input_pdf, "wb") as f:
            f.write(uploaded_file.getbuffer())

        output_docx = temp_path / f"{filename_base}.docx"

        status_text.text(f"Processing {idx+1}/{len(uploaded_files)}: {uploaded_file.name}")

        try:
            cv = Converter(str(input_pdf))

            if convert_all:
                cv.convert(str(output_docx))
            else:
                cv.convert(str(output_docx), start=start_page-1, end=end_page)

            cv.close()

            # Post-process the DOCX for perfect Arabic
            from docx import Document
            doc = Document(str(output_docx))

            for para in doc.paragraphs:
                if para.text.strip():
                    # Fix Arabic text shaping and direction
                    reshaped_text = arabic_reshaper.reshape(para.text)
                    bidi_text = get_display(reshaped_text)

                    para.text = bidi_text
                    para.paragraph_format.right_to_left = True
                    para.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                    for run in para.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(12)
                        run.font.bold = False  # Optional: adjust as needed

            # Save improved version
            doc.save(str(output_docx))

            output_files.append((output_docx, f"{filename_base}.docx"))

        except Exception as e:
            st.error(f"Failed to convert {uploaded_file.name}: {str(e)}")
            continue

        progress_bar.progress((idx + 1) / len(uploaded_files))

    status_text.text("All done!")
    progress_bar.empty()

    # Download
    if len(output_files) == 1:
        with open(output_files[0][0], "rb") as f:
            st.download_button(
                label="📥 Download Word File",
                data=f,
                file_name=output_files[0][1],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        zip_path = temp_path / "arabic_converted_files.zip"
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for docx_path, arcname in output_files:
                zipf.write(docx_path, arcname)
        with open(zip_path, "rb") as f:
            st.download_button(
                label=f"📦 Download all {len(output_files)} files as ZIP",
                data=f,
                file_name="arabic_converted_files.zip",
                mime="application/zip"
            )

    st.success("Conversion complete! Arabic text is now properly formatted.")
    st.balloons()

st.markdown("---")
st.caption("Made with ❤️ for perfect Arabic document conversion")
