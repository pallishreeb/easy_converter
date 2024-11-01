import os
from flask import Flask, request, render_template, send_file, redirect
from pdf2docx import Converter
from werkzeug.utils import secure_filename
from PIL import Image
from rembg import remove
import io
from PyPDF2 import PdfReader, PdfWriter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024 # 10 MB limit
app.config['UPLOAD_FOLDER'] = 'uploads'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def upload_form():
    return render_template('upload.html')
    
@app.route('/resize', methods=['GET'])
def resize_image_form():
    return render_template('resize.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)

    if file and file.filename.endswith('.pdf'):
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)

        # Convert PDF to Word
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename.replace('.pdf', '.docx'))
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()

        # Send the file to the user for download
        response = send_file(docx_path, as_attachment=True)

        # Remove the files after download
        os.remove(pdf_path)  # Remove the PDF file
        os.remove(docx_path)  # Remove the DOCX file

        return response

    return 'Invalid file format. Please upload a PDF.'

def preprocess_image(image):
    # Convert image to RGB (if not already in RGB)
    if image.mode != 'RGB':
        image = image.convert('RGB')

    # Enhance image contrast (optional)
    from PIL import ImageEnhance
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(1.5)  # Increase contrast by 50%

    return image

@app.route('/resize_image', methods=['POST'])
def resize_image():
    file = request.files['file']
    if file and file.filename.endswith(('.png', '.jpg', '.jpeg')):
        img = Image.open(file)

        # Original size
        original_size = img.size
        original_format = img.format

        # Handle resizing
        resolution = request.form.get('resolution')
        if resolution == 'low':
            img = img.resize((int(original_size[0] * 0.5), int(original_size[1] * 0.5)))
        elif resolution == 'medium':
            img = img.resize((int(original_size[0] * 0.75), int(original_size[1] * 0.75)))
        elif resolution == 'high':
            img = img.resize((original_size[0], original_size[1]))

        # Handle background removal
        if 'background_removal' in request.form:
            img_data = io.BytesIO()
            img.save(img_data, format=original_format)
            img_data.seek(0)

            try:
                img_bytes = img_data.read()
                img_removed_bg = remove(img_bytes)
                img = Image.open(io.BytesIO(img_removed_bg))
            except Exception as e:
                return f"Error removing background: {str(e)}"

        # Handle format conversion
        format_choice = request.form.get('format')
        if format_choice not in ['png', 'jpeg']:
            return 'Invalid format choice. Please choose either PNG or JPEG.'

        # Convert to RGB if saving as JPEG
        if format_choice == 'jpeg' and img.mode == 'RGBA':
            img = img.convert('RGB')

        output_path = f"output_image.{format_choice}"
        img.save(output_path, format=format_choice.upper())  # Save with the correct format

        # Send the file
        response = send_file(output_path, as_attachment=True)

        # Delete the output file after sending it
        os.remove(output_path)

        return response


    return 'Invalid file format. Please upload an image.'
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'png', 'jpg', 'jpeg', 'gif'}

@app.route('/about')
def about():
    return render_template('about.html')  # Assuming you save the above HTML as about.html
@app.route('/split_pdf')
def split_pdf_form():
    return render_template('split.html') 

@app.route('/upload_pdf', methods=['POST'])
def upload_pdf():
    file = request.files['file']
    if file and file.filename.endswith('.pdf'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        return send_file(filepath, as_attachment=False)

    return 'Invalid file format. Please upload a PDF.'
    
@app.route('/extract_pdf_pages', methods=['POST'])
def extract_pdf_pages():
    print("Request Files:", request.files)  # Log the request files
    print("Request Form:", request.form)    # Log the request form data

    # Get the uploaded file
    file = request.files.get('file')
    if file is None:
        return "No file uploaded.", 400

    print("Uploaded File Name:", file.filename)  # Log the filename

    # Check if the uploaded file is a PDF
    if file.filename.endswith('.pdf'):
        try:
            # Read start and end pages from the form
            start_page = int(request.form.get('start_page'))
            end_page = int(request.form.get('end_page'))

            # Read the PDF file
            reader = PdfReader(file)
            writer = PdfWriter()

            # Ensure that the page numbers are within valid range
            if start_page < 1 or end_page < 1 or start_page > len(reader.pages) or end_page > len(reader.pages):
                return "Invalid page range. Please ensure the pages are within the PDF limits.", 400
            
            # Adjust for zero-based index and extract pages
            for i in range(start_page - 1, end_page):
                writer.add_page(reader.pages[i])

            # Create a BytesIO stream for the output PDF
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)

            # Send the file as a response
            response = send_file(output, as_attachment=True, download_name="extracted_pages.pdf", mimetype='application/pdf')

            # No need to delete a physical file as we are using BytesIO
            
            return response

        except Exception as e:
            return f"Error processing PDF: {str(e)}", 500
    else:
        return "Invalid file format. Please upload a PDF.", 400
if __name__ == '__main__':
    app.run(debug=True)
