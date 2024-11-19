
# Multi-language Document Translator

This is a Streamlit-based web application for multi-language document translation. It supports text extraction from PDFs cleans and structures the extracted text using Google's Gemini AI, and translates the text into a target language of choice.

---

## Features

- **Multi-language OCR**: Supports text extraction from PDFs in various languages, including English, German, French, Spanish, Hindi, Tamil, Telugu, Malayalam, Kannada, Marathi, and Russian.
- **AI-powered Cleaning**: Utilizes Google's Gemini AI to clean and structure extracted text, preserving document format while removing irrelevant content.
- **Translation**: Provides translation to any supported language using AI.
- **Export Options**: Allows users to download the translated text as a PDF document.
- **Interactive UI**: Simple and user-friendly interface powered by Streamlit.

---

## Tech Stack

- **Programming Language**: Python
- **Framework**: Streamlit
- **Libraries**:
  - `fitz` (PyMuPDF) for PDF text extraction
  - `pytesseract` for OCR
  - `google.generativeai` for AI-powered text cleaning and translation
  - `docx` and `pythoncom` for Word and PDF file handling
  - `langdetect` for language detection
  - `pdf2image` for converting PDFs to images for OCR

---

## Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/your-username/multi-language-document-translator.git
   cd multi-language-document-translator
   ```

2. **Set Up Environment**
   Ensure you have Python installed (>= 3.7). Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate    # On Windows: venv\Scripts\activate

4. **Set Up External Tools**
   - Install [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) and set the `pytesseract.pytesseract.tesseract_cmd` path in the code.
   - Install [Poppler](http://blog.alivate.com.au/poppler-windows/) and set the `POPPLER_PATH`.

5. **Configure Gemini API**
   Obtain your API key for Google's Gemini AI and set it in the code:
   ```python
   gemini_key = "YOUR_GEMINI_API_KEY"
   ```

---

## Usage

1. **Run the Application**
   ```bash
   streamlit run app.py
   ```

2. **Upload Document**
   - Upload a PDF file for text extraction.
   - Choose the input language for OCR.

3. **Clean and Translate**
   - View the extracted text.
   - Clean and structure the text using AI.
   - Translate the text to the target language of your choice.

4. **Download Translated Document**
   - Download the translated text as a PDF document.

---

## Supported Languages

- English
- German
- French
- Spanish
- Hindi
- Tamil
- Telugu
- Malayalam
- Kannada
- Marathi
- Russian

---

## Troubleshooting

- **No text extracted**: Ensure the uploaded document has readable text or consider using OCR for image-based PDFs.
- **OCR not working**: Verify `Tesseract` and `Poppler` are correctly installed and their paths are configured.
- **Gemini AI errors**: Ensure your API key is valid and the Gemini API service is accessible.


