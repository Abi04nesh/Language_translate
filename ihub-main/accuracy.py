import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
import pytesseract
from pdf2image import convert_from_path
from ocr_tamil.ocr import OCR
import os
import tempfile
from PIL import Image
from docx import Document
from langdetect import detect
from docx2pdf import convert  # For converting DOCX to PDF
from difflib import SequenceMatcher  # For accuracy comparison
import pythoncom  # For COM initialization
from win32com.client import Dispatch  # For COM controls
from io import BytesIO
import logging  # For debugging

# print the integration level
logging.basicConfig(level=logging.INFO)

gemini_key = "AIzaSyCv_e5Ozpdf4FfUWh8b9_dDEye0y-7t0jU"
genai.configure(api_key=gemini_key)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\Program Files\poppler-24.08.0\Library\bin'

OCR_LANGUAGES = {
    "English": "eng",
    "Hindi": "hin",
    "Telugu": "tel",
    "Tamil": "ta",
    "Malayalam": "mal",
    "Kannada": "kan",
    "Marathi": "mar",
    "Russian": "rus",
    "German": "deu",
    "French": "fra",
    "Spanish": "spa"
}

def detect_language(text):
    try:
        return detect(text)
    except Exception as e:
        st.error(f"Error detecting language: {e}")
        return None

# Extract text with PyMuPDF (fallback to OCR if PyMuPDF fails)
def extract_text_with_fallback(uploaded_file, ocr_language):
    try:
        logging.info("Starting PDF text extraction using PyMuPDF.")
        pdf_data = BytesIO(uploaded_file.read())

        with fitz.open(stream=pdf_data, filetype="pdf") as doc:
            text = ""
            for page_num, page in enumerate(doc, start=1):
                logging.info(f"Extracting text from page {page_num}.")
                page_text = page.get_text()
                if page_text:
                    text += page_text
            if text.strip():
                logging.info("Text successfully extracted using PyMuPDF.")
                return text.strip()

        logging.warning("PyMuPDF failed. Attempting OCR fallback.")
        temp_pdf_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        images = convert_from_path(temp_pdf_path, poppler_path=POPPLER_PATH)
        ocr_text = ""
        for page_num, image in enumerate(images, start=1):
            logging.info(f"Performing OCR on page {page_num}.")
            ocr_text += pytesseract.image_to_string(image, lang=ocr_language)

        if ocr_text.strip():
            logging.info("Text successfully extracted using OCR.")
            return ocr_text.strip()

        logging.error("No text could be extracted using PyMuPDF or OCR.")
        st.error("No text could be extracted from the document.")
        return None

    except Exception as e:
        logging.error(f"Error during text extraction: {e}")
        st.error(f"Error extracting text: {e}")
        return None

# Clean and structure extracted text using Gemini AI
def clean_and_structure_text(extracted_text):
    try:
        logging.info("Cleaning and structuring extracted text using Gemini AI.")
        prompt_clean_text = (
            "Analyze and structure the following text extracted from a document, keeping the headings and body of the text intact. Remove the things like images and irrelevant strings:\n\n"
            + extracted_text
        )
        response_clean_text = model.generate_content(prompt_clean_text)
        return response_clean_text.text.strip()
    except Exception as e:
        logging.error(f"Error during text cleaning: {e}")
        st.error(f"Error cleaning text: {e}")
        return None

# Save cleaned text to a Word document
def save_text_to_word(text, file_path):
    try:
        logging.info(f"Saving text to Word document at {file_path}.")
        doc = Document()
        doc.add_paragraph(text)
        doc.save(file_path)
    except Exception as e:
        logging.error(f"Error saving Word document: {e}")
        st.error(f"Error saving Word document: {e}")

# Convert DOCX to PDF
def convert_docx_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()
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

# Accuracy Checker (contextual similarity)
def calculate_translation_accuracy(input_text, translated_text):
    try:
        logging.info("Calculating translation accuracy.")
        ratio = SequenceMatcher(None, input_text, translated_text).ratio()
        accuracy = round(ratio * 100, 2)
        return accuracy
    except Exception as e:
        logging.error(f"Error calculating accuracy: {e}")
        st.error(f"Error calculating accuracy: {e}")
        return None

model = genai.GenerativeModel('models/gemini-1.5-pro')

# Streamlit web app
st.title("Multi-language Document Translator ")

input_language = st.selectbox("Select the document language:", OCR_LANGUAGES.keys())
ocr_language = OCR_LANGUAGES[input_language]

uploaded_file = st.file_uploader("Upload a document (PDF or DOCX):", type=["pdf", "docx"])

if uploaded_file:
    with st.spinner('Extracting text from document...'):
        extracted_text = extract_text_with_fallback(uploaded_file, ocr_language)
        if extracted_text:
            logging.info("Text extraction complete.")
            recognized_language = detect_language(extracted_text)
            st.write(f"Detected Language: {recognized_language.capitalize()}")
            cleaned_text = clean_and_structure_text(extracted_text)
            st.subheader("Cleaned and Structured Text:")
            st.text_area("Cleaned Text:", value=cleaned_text, height=300, key="cleaned_text")
        else:
            st.error("No text extracted from the uploaded document.")

# Select target language for translation
target_language = st.selectbox("Select the language to Translate:", OCR_LANGUAGES.keys())

# Translate button
if st.button("Translate"):
    cleaned_text = st.session_state.get("cleaned_text")
    if cleaned_text:
        with st.spinner(f'Translating to {target_language}...'):
            prompt_to_target = (
                f"Translate the following text to {target_language}:\n\n" + cleaned_text
            )
            response_to_target = model.generate_content(prompt_to_target)
            target_text = response_to_target.text.strip()
            st.subheader(f"Translated {target_language} Text:")
            st.text_area(f"Translated {target_language} Text:", value=target_text, height=300, key="target_text")

            st.session_state["translated_text"] = target_text
    else:
        st.error("Please clean and structure the text first.")

# Accuracy check button
if st.button("Check Accuracy"):
    input_text = st.session_state.get("cleaned_text")
    translated_text = st.session_state.get("translated_text")
    if input_text and translated_text:
        accuracy = calculate_translation_accuracy(input_text, translated_text)
        st.write(f"Translation Accuracy: {accuracy}%")
    else:
        st.error("Please ensure both input and translated text are available.")

# Download PDF button
if st.session_state.get("translated_text"):
    if st.button("Download Translated Document as PDF"):
        target_text = st.session_state.get("translated_text")
        if target_text:
            with st.spinner('Generating PDF document...'):
                word_file_path = os.path.join(tempfile.gettempdir(), "translated_docs.docx")
                save_text_to_word(target_text, word_file_path)

                pdf_file_path = os.path.join(tempfile.gettempdir(), "translated_docs.pdf")
                convert_docx_to_pdf(word_file_path, pdf_file_path)

                with open(pdf_file_path, "rb") as file:
                    st.download_button(label="Download PDF", data=file, file_name="translated_docs.pdf", mime="application/pdf")
        else:
            st.error("Please translate the text first.")
else:
    st.write("Perform translation to enable download.")
