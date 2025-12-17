import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import zipfile
from pathlib import Path

# Page config & styling
st.set_page_config(page_title="Arabic PDF → Word Converter", page_icon="📄")
st.title("🇸🇦 Arabic PDF to Word Converter")
st.markdown("""
Upload one or more Arabic PDFs and convert them to editable Word (.docx) files.  
Layout, tables, images, and right-to-left text are preserved as much as possible.
""")

# Sidebar options
st.sidebar.header("Conversion Options")
convert_all_pages = st.sidebar.checkbox("Convert all pages", value=True)

if not convert_all_pages:
    col1, col2 = st.sidebar.columns(2)
    start_page = col1.number_input("Start page", min_value=1, value=1, step=1)
    end_page = col2.number_input("End page", min_value=1, value=10, step=1)
    if start_page > end_page:
        st.sidebar.error("Start page must be ≤ End page")
        st.stop()
else:
    start_page = None
    end_page = None

# File uploader – allow multiple
uploaded_files = st.file_uploader(
    "Choose PDF file(s)",
    type="pdf",
    accept_multiple_files=True,
    help="You can select multiple PDFs at once"
)

if not uploaded_files:
    st.info("👆 Upload one or more PDF files to get started.")
    st.stop()

# Temporary directory for all processing
with tempfile.TemporaryDirectory() as temp_dir:
    temp_path = Path(temp_dir)
    output_files = []  # List to collect all .docx paths

    # Progress bar setup
    total_tasks = len(uploaded_files)
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uploaded_file in enumerate(uploaded_files):
        filename_base = Path(uploaded_file.name).stem

        # Save uploaded PDF
        input_pdf_path = temp_path / uploaded_file.name
        with open(input_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Output Word path
        output_docx_path = temp_path / f"{filename_base}.docx"

        # Update status
        status_text.text(f"Processing {idx+1}/{total_tasks}: {uploaded_file.name}")

        try:
            cv = Converter(str(input_pdf_path))

            # Get total pages for better progress
            cv.load_pages()  # Needed to access page count
            total_pages = len(cv.pages)
            page_range_text = "all pages" if convert_all_pages else f"pages {start_page}–{end_page}"

            # Inner progress for pages
            page_progress = st.empty()

            # Custom callback to update page progress
            def page_callback(current_page: int):
                progress_percent = current_page / total_pages
                page_progress.progress(progress_percent)
                page_progress.text(f"Converting page {current_page} of {total_pages}")

            # Convert with page range
            if convert_all_pages:
                cv.convert(str(output_docx_path), callback=page_callback)
            else:
                # pdf2docx uses 0-based indexing
                cv.convert(str(output_docx_path), start=start_page-1, end=end_page, callback=page_callback)

            cv.close()
            page_progress.empty()  # Clear page progress

            output_files.append((output_docx_path, f"{filename_base}.docx"))

        except Exception as e:
            st.error(f"Failed to convert **{uploaded_file.name}**: {str(e)}")
            continue

        # Update overall progress
        progress_bar.progress((idx + 1) / total_tasks)

    # All done
    status_text.text("All conversions completed!")
    progress_bar.empty()

    # Provide downloads
    if len(output_files) == 1:
        # Single file → direct download
        docx_path, display_name = output_files[0]
        with open(docx_path, "rb") as f:
            st.download_button(
                label="📥 Download Word file",
                data=f,
                file_name=display_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        # Multiple files → ZIP
        zip_path = temp_path / "converted_word_files.zip"
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for docx_path, arcname in output_files:
                zipf.write(docx_path, arcname)

        with open(zip_path, "rb") as f:
            st.download_button(
                label=f"📦 Download all {len(output_files)} Word files as ZIP",
                data=f,
                file_name="arabic_converted_files.zip",
                mime="application/zip"
            )

    st.success("Done! Your files are ready.")
    st.balloons()
