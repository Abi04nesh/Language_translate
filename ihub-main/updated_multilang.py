import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
import pytesseract
from pdf2image import convert_from_path
import os
import tempfile
from PIL import Image
from docx import Document
from langdetect import detect
from docx2pdf import convert  # For converting DOCX to PDF
import pythoncom  # For COM initialization
from win32com.client import Dispatch  # For COM controls
from io import BytesIO
import logging  # For debugging

# Initialize logging
logging.basicConfig(level=logging.INFO)

gemini_key = "AIzaSyDJdhKPF0yMY3q6MMYpxAoKFdKhkgbX6U0"
genai.configure(api_key=gemini_key)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\Program Files\poppler-24.08.0\Library\bin'


# Supported languages for OCR
OCR_LANGUAGES = {
    "English": "eng",
    "German": "deu",
    "French": "fra",
    "Spanish": "spa",
    "Hindi": "hin",
    "Tamil": "tam",
    "Telugu": "tel",
    "Malayalam": "mal",
    "Kannada": "kan",
    "Marathi": "mar",
    "Russian": "rus"
}

# Function to detect language of extracted text
def detect_language(text):
    try:
        return detect(text)
    except Exception as e:
        st.error(f"Error detecting language: {e}")
        return None

# Function to extract text using PyMuPDF or OCR as fallback
def extract_text_with_fallback(uploaded_file, ocr_language):
    try:
        logging.info("Starting document text extraction.")
        pdf_data = BytesIO(uploaded_file.read())

        # Attempt to extract text using PyMuPDF
        with fitz.open(stream=pdf_data, filetype="pdf") as doc:
            text = ""
            for page_num in range(len(doc)):
                logging.info(f"Extracting text from page {page_num + 1}.")
                page_text = doc[page_num].get_text()
                text += page_text

            if text.strip():
                logging.info("Text successfully extracted using PyMuPDF.")
                return text.strip()

        logging.warning("PyMuPDF failed to extract text. Attempting OCR fallback...")

        # Save the uploaded PDF to a temp file
        temp_pdf_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Convert PDF to images for OCR
        images = convert_from_path(temp_pdf_path, poppler_path=POPPLER_PATH)

        ocr_text = ""
        for page_num, image in enumerate(images, start=1):
            logging.info(f"Performing OCR on page {page_num}.")
            ocr_text += pytesseract.image_to_string(image, lang=ocr_language)

        if ocr_text.strip():
            logging.info("Text successfully extracted using OCR.")
            return ocr_text.strip()

        logging.error("No text could be extracted from the document.")
        st.error("No text could be extracted from the document.")
        return None

    except Exception as e:
        logging.error(f"Error during text extraction: {e}")
        st.error(f"Error extracting text: {e}")
        return None

# Function to clean and structure extracted text using Gemini AI
def clean_and_structure_text(extracted_text):
    try:
        logging.info("Cleaning and structuring extracted text using Gemini AI.")
        prompt_clean_text = (
            " Preserve the full structure of the text extracted from a document, keeping the format exactly as per document, while keeping the headings and full body of the text intact. Remove the images and irrelevant hyperlinks:\n\n"
            + extracted_text
        )
        response_clean_text = model.generate_content(prompt_clean_text)
        return response_clean_text.text.strip()
    except Exception as e:
        logging.error(f"Error during cleaning and structuring text: {e}")
        st.error(f"Error cleaning and structuring text: {e}")
        return None

# Function to save text to a Word document
def save_text_to_word(text, file_path):
    try:
        logging.info(f"Saving text to Word document at {file_path}.")
        doc = Document()
        doc.add_paragraph(text)
        doc.save(file_path)
    except Exception as e:
        logging.error(f"Error saving text to Word document: {e}")
        st.error(f"Error saving text to Word document: {e}")

# Function to convert DOCX to PDF with COM initialization
def convert_docx_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()  # Initialize COM libraries
    try:
        logging.info(f"Converting DOCX to PDF: {docx_path} -> {pdf_path}.")
        word = Dispatch("Word.Application")
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
    except Exception as e:
        logging.error(f"Failed to convert DOCX to PDF: {e}")
        st.error(f"Failed to convert DOCX to PDF: {e}")
    finally:
        pythoncom.CoUninitialize()

# Initialize the generative model
model = genai.GenerativeModel('models/gemini-1.5-flash')

st.title("Multi-language Document Translator")

# Language selection for OCR (Optical character recognition)
input_language = st.selectbox("Select the input language for the document:", OCR_LANGUAGES.keys())
ocr_language = OCR_LANGUAGES[input_language]

# File uploader for PDF or document
uploaded_file = st.file_uploader("Upload a document (PDF):", type=["pdf"])

if uploaded_file:
    with st.spinner('Extracting text from the document...'):
        extracted_text = extract_text_with_fallback(uploaded_file, ocr_language)
        if extracted_text:
            logging.info("Text extraction complete.")
            recognized_language = detect_language(extracted_text)
            st.write(f"Detected Language: {recognized_language.capitalize()}")
            cleaned_text = clean_and_structure_text(extracted_text)
            st.subheader("Cleaned and Structured Text:")
            st.text_area("Cleaned Text:", cleaned_text, height=300, key="cleaned_text")
        else:
            st.error("No text extracted from the uploaded document.")

# Language selection for translation
target_language = st.selectbox("Select the target language:", list(OCR_LANGUAGES.keys()))

# Button to generate translated text
if st.button("Translate"):
    cleaned_text = st.session_state.get("cleaned_text")
    if cleaned_text:
        with st.spinner(f'Translating to {target_language}...'):
            prompt_to_target = (
                f"Translate the extracted {input_language} text fully to {target_language}:\n\n" + cleaned_text
            )

            response_to_target = model.generate_content(prompt_to_target)
            target_text = response_to_target.text.strip()

            st.subheader(f"Translated {target_language} Text:")
            st.text_area(f"Translated {target_language} Text:", target_text, height=300, key="target_text")

            st.session_state["translated_text"] = target_text
    else:
        st.error("Please clean and structure the text first.")

if st.session_state.get("translated_text"):
    if st.button("Download Translated Document as PDF"):
        target_text = st.session_state.get("translated_text")
        if target_text:
            with st.spinner('Generating PDF document...'):
                
                word_file_path = os.path.join(tempfile.gettempdir(), "translated_doc.docx")
                save_text_to_word(target_text, word_file_path)

                pdf_file_path = os.path.join(tempfile.gettempdir(), "translated_doc.pdf")
                convert_docx_to_pdf(word_file_path, pdf_file_path)

                # Allow the user to download the PDF file
                with open(pdf_file_path, "rb") as file:
                    st.download_button(label="Download PDF Document", data=file, file_name="translated_doc.pdf", mime="application/pdf")
        else:
            st.error("Please translate the text first.")
else:
    st.write("Complete the translation to enable download.")
