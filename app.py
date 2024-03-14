from flask import Flask, render_template, request, send_file, make_response
import os
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
from google.cloud import vision
from io import BytesIO
from PIL import Image
from docx import Document

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
WORD_FOLDER = 'word_documents'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(WORD_FOLDER):
    os.makedirs(WORD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['WORD_FOLDER'] = WORD_FOLDER

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "test.json"  # Replace with your service account key

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        file = request.files['file']

        if file and (file.filename.endswith('.pdf') or file.filename.endswith('.jpg') or file.filename.endswith('.jpeg') or file.filename.endswith('.png')):
            pdf_filename = secure_filename(file.filename)
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
            file.save(pdf_path)

            if file.filename.endswith('.pdf'):
                # Convert PDF pages to images
                images = convert_pdf_to_images(pdf_path)
            else:
                images = [Image.open(pdf_path)]

            # Process each page image using Google Cloud Vision
            detected_texts = []
            client = vision.ImageAnnotatorClient()
            for i, image in enumerate(images):
                # Process the image content using Google Cloud Vision
                image_content = BytesIO()
                image.save(image_content, format='PNG')  # Convert to PNG format before processing
                image_content.seek(0)
                vision_image = vision.Image(content=image_content.read())
                response = client.text_detection(image=vision_image)
                texts = response.text_annotations

                if texts:
                    detected_text = texts[0].description
                    detected_texts.append((i + 1, detected_text))  # Page numbers start from 1

            # Save extracted text to Word document
            word_filename = save_text_to_word(pdf_filename, detected_texts, app.config['WORD_FOLDER'])
            
            return render_template('result.html', pdf_filename=pdf_filename, detected_texts=detected_texts, word_filename=word_filename)

    return 'No PDF or image file detected', 400

def convert_pdf_to_images(pdf_path):
    images = []
    pdf_document = fitz.open(pdf_path)
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        # Convert PDF page to image format (PNG)
        pix = page.get_pixmap()
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(image)
    pdf_document.close()
    return images

def save_text_to_word(pdf_filename, detected_texts, word_folder):
    word_document = Document()
    for page, text in detected_texts:
        word_document.add_paragraph(f'Page {page}:\n{text}\n')
    word_filename = f"{os.path.splitext(pdf_filename)[0]}.docx"
    word_path = os.path.join(word_folder, word_filename)
    word_document.save(word_path)
    return word_filename

def save_text_to_pdf(pdf_filename, detected_texts, pdf_folder):
    pdf_filename = os.path.splitext(pdf_filename)[0] + "_text.pdf"
    pdf_path = os.path.join(pdf_folder, pdf_filename)
    with open(pdf_path, 'w') as f:
        for page, text in detected_texts:
            f.write(f'Page {page}:\n{text}\n')
            f.write('\n')
    return pdf_filename

@app.route('/download_word/<filename>')
def download_word(filename):
    word_path = os.path.join(app.config['WORD_FOLDER'], filename)
    if os.path.exists(word_path):
        return send_file(word_path, as_attachment=True)
    else:
        return 'Word file not found', 404

@app.route('/download_text/<filename>')
def download_text(filename):
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(pdf_path):
        # Convert PDF pages to images
        images = convert_pdf_to_images(pdf_path)

        # Process each page image using Google Cloud Vision
        detected_texts = []
        client = vision.ImageAnnotatorClient()
        for i, image in enumerate(images):
            # Process the image content using Google Cloud Vision
            image_content = BytesIO()
            image.save(image_content, format='PNG')  # Convert to PNG format before processing
            image_content.seek(0)
            vision_image = vision.Image(content=image_content.read())
            response = client.text_detection(image=vision_image)
            texts = response.text_annotations

            if texts:
                detected_text = texts[0].description
                detected_texts.append((i + 1, detected_text))  # Page numbers start from 1

        # Save extracted text to PDF
        pdf_filename = save_text_to_pdf(filename, detected_texts, app.config['UPLOAD_FOLDER'])
        
        return send_file(os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename), as_attachment=True)
    else:
        return 'PDF file not found', 404

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename))

if __name__ == '__main__':
    app.run(debug=True)
