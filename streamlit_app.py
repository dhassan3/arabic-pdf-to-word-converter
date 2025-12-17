import streamlit as st
from pdf2image import convert_from_bytes
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Inches
from pathlib import Path
import tempfile
import zipfile
from io import BytesIO

st.set_page_config(page_title="Arabic PDF → Word", page_icon="📄")
st.title("🇸🇦 Arabic PDF to Word Converter")
st.markdown("Convert Arabic PDFs to editable Word files. Great for text and scanned documents.")

# Sidebar
st.sidebar.header("Options")
all_pages = st.sidebar.checkbox("Convert all pages", value=True)
if not all_pages:
    col1, col2 = st.sidebar.columns(2)
    start = col1.number_input("Start page", min_value=1, value=1, step=1)
    end = col2.number_input("End page", min_value=1, value=10, step=1)
    if start > end:
        st.sidebar.error("Start must be ≤ End")
        st.stop()

image_mode = st.sidebar.checkbox(
    "Exact layout (embed as images)",
    value=False,
    help="Best for scanned PDFs – perfect visual match (text not searchable)"
)

# Upload
files = st.file_uploader("Upload PDF(s)", type="pdf", accept_multiple_files=True)
if not files:
    st.info("Upload PDFs to start.")
    st.stop()

with tempfile.TemporaryDirectory() as tmp:
    tmp_path = Path(tmp)
    results = []
    progress = st.progress(0)
    status = st.empty()

    for i, file in enumerate(files):
        base = Path(file.name).stem
        pdf_bytes = file.getvalue()

        # Preview
        if i == 0 or len(files) == 1:
            with st.expander("👁️ View uploaded PDF"):
                st.download_button("Open locally", pdf_bytes, file.name, "application/pdf")

        # Scanned detection
        reader = PdfReader(BytesIO(pdf_bytes))
        text_len = sum(len(page.extract_text() or "") for page in reader.pages)
        if text_len < 100:
            st.warning(f"⚠️ {file.name} seems scanned. Use 'Exact layout' mode for best results.")

        docx_path = tmp_path / f"{base}.docx"
        status.text(f"Processing {i+1}/{len(files)}: {file.name}")

        try:
            word = Document()
            word.add_heading(f"From: {file.name}", 1)

            total = len(reader.pages)
            page_bar = st.empty()

            page_range = range(total) if all_pages else range(start-1, end)

            for n in page_range:
                page_bar.progress((n - (start-1 if not all_pages else 0) + 1) / len(page_range))
                page_bar.text(f"Page {n+1}/{total}")

                if image_mode:
                    # High-quality image
                    images = convert_from_bytes(pdf_bytes, dpi=300, first_page=n+1, last_page=n+1)
                    img_bytes = BytesIO()
                    images[0].save(img_bytes, format="PNG")
                    img_bytes.seek(0)
                    word.add_paragraph().add_run().add_picture(img_bytes, width=Inches(6.5))
                    word.add_page_break()
                else:
                    # Text extraction
                    page_text = reader.pages[n].extract_text() or ""
                    if page_text.strip():
                        p = word.add_paragraph(page_text)
                        if any("\u0600" <= c <= "\u06FF" for c in page_text):
                            p.paragraph_format.right_to_left = True

            word.save(str(docx_path))
            page_bar.empty()
            results.append((docx_path, f"{base}.docx"))

        except Exception as e:
            st.error(f"Error with {file.name}: {e}")

        progress.progress((i+1)/len(files))

    status.text("Complete!")
    progress.empty()

    # Download
    if len(results) == 1:
        with open(results[0][0], "rb") as f:
            st.download_button("📥 Download Word file", f, file_name=results[0][1],
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        zip_path = tmp_path / "converted.zip"
        with zipfile.ZipFile(zip_path, "w") as z:
            for p, n in results:
                z.write(p, n)
        with open(zip_path, "rb") as f:
            st.download_button("📦 Download all as ZIP", f, "arabic_converted.zip", "application/zip")

    st.success("Success!")
    st.balloons()

    st.caption("Was this helpful?")
    st.feedback("thumbs")

st.markdown("Made with ❤️ for Arabic documents")
