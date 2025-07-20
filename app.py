from flask import Flask, render_template, request, send_file
from pdf2image import convert_from_path
from pdf2docx import Converter
import os
import zipfile

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    pdf_file = request.files['pdf_file']
    output_format = request.form['output_format']
    input_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
    pdf_file.save(input_path)

    if output_format == 'jpeg':
        pages = convert_from_path(input_path)
        image_paths = []
        for i, page in enumerate(pages):
            image_path = os.path.join(CONVERTED_FOLDER, f'page_{i+1}.jpg')
            page.save(image_path, 'JPEG')
            image_paths.append(image_path)

        if len(image_paths) > 1:
            zip_path = os.path.join(CONVERTED_FOLDER, 'images.zip')
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for image in image_paths:
                    zipf.write(image, os.path.basename(image))
            return send_file(zip_path, as_attachment=True)
        else:
            return send_file(image_paths[0], as_attachment=True)

    elif output_format == 'word':
        docx_path = os.path.join(CONVERTED_FOLDER, 'output.docx')
        cv = Converter(input_path)
        cv.convert(docx_path)
        cv.close()
        return send_file(docx_path, as_attachment=True)

    else:
        return 'Invalid format selected'

if __name__ == '__main__':
    app.run(debug=True)
