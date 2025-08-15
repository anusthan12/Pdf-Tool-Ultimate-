# ocr_backend.py
import os
import io
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename

# OCR and PDF processing libraries
import fitz  # PyMuPDF
from pdf2image import convert_from_bytes
import pytesseract
from PIL import Image
from docx import Document

# --- App Initialization ---
app = Flask(__name__)
CORS(app) # Allow requests from the frontend domain

# --- Helper Functions ---
def convert_without_ocr(pdf_bytes):
    """Extracts text directly from a PDF and returns a Word document."""
    doc = Document()
    doc.add_heading('Converted Document (Text Extraction)', 0)
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        doc.add_paragraph(text)
        if page_num < len(pdf_document) - 1:
            doc.add_page_break()
    return save_doc_to_memory(doc)

def convert_with_ocr(pdf_bytes):
    """Converts PDF pages to images and uses Tesseract OCR to extract text."""
    doc = Document()
    doc.add_heading('Converted Document (OCR)', 0)
    images = convert_from_bytes(pdf_bytes)
    for i, image in enumerate(images):
        text = pytesseract.image_to_string(image)
        doc.add_paragraph(text)
        if i < len(images) - 1:
            doc.add_page_break()
    return save_doc_to_memory(doc)

def save_doc_to_memory(doc):
    """Saves a docx.Document object to an in-memory bytes buffer."""
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- API Endpoint ---
@app.route('/convert-pdf-to-word', methods=['POST'])
def handle_conversion():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    mode = request.form.get('mode', 'no_ocr')

    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and file.filename.lower().endswith('.pdf'):
        pdf_bytes = file.read()
        try:
            if mode == 'ocr':
                buffer = convert_with_ocr(pdf_bytes)
            else: # Default to 'no_ocr'
                buffer = convert_without_ocr(pdf_bytes)
            
            return send_file(
                buffer,
                as_attachment=True,
                download_name=f'converted_{mode}.docx',
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        except Exception as e:
            return jsonify({"error": f"An error occurred during processing: {str(e)}"}), 500
    else:
        return jsonify({"error": "Invalid file type. Please upload a PDF."}), 400

# --- Health Check Route ---
@app.route('/')
def index():
    return "OCR Backend is running!", 200

# --- For local testing ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
