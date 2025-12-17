import streamlit as st
from pdf2docx import Converter
import os
import tempfile

# App title and description
st.title("Arabic PDF to Word Converter")
st.markdown("""
This app converts Arabic PDFs to Word (.docx) files while preserving text, layout, and RTL direction.  
Upload your PDF below and download the converted file.
""")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    # Create a temporary directory to save files
    with tempfile.TemporaryDirectory() as temp_dir:
        input_pdf_path = os.path.join(temp_dir, uploaded_file.name)
        
        # Save uploaded PDF to temp path
        with open(input_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Define output path
        output_docx_path = os.path.join(temp_dir, f"{os.path.splitext(uploaded_file.name)[0]}.docx")
        
        # Progress indicator
        with st.spinner("Converting PDF to Word..."):
            try:
                # Initialize converter
                cv = Converter(input_pdf_path)
                
                # Convert all pages
                cv.convert(output_docx_path)
                
                # Close converter
                cv.close()
                
                st.success("Conversion complete!")
                
                # Download button
                with open(output_docx_path, "rb") as f:
                    st.download_button(
                        label="Download Word File",
                        data=f,
                        file_name=os.path.basename(output_docx_path),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
else:
    st.info("Please upload a PDF file to start.")
