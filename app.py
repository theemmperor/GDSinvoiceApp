from flask import Flask, render_template, request, jsonify, send_from_directory, redirect, url_for, flash
from utils.screw_press_handler import ScrewPressHandler
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.secret_key = 'your-secret-key-here'  # Required for flash messages

# Configure upload folders
UPLOAD_FOLDER = 'data'
IMAGE_FOLDER = os.path.expanduser('~/Desktop/invoice_images')
LOGO_FOLDER = os.path.join(IMAGE_FOLDER, 'logos')

# Create required directories if they don't exist
for folder in [UPLOAD_FOLDER, IMAGE_FOLDER, LOGO_FOLDER]:
    os.makedirs(folder, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('upload.html', error='No file selected')

        file = request.files['file']
        if file.filename == '':
            return render_template('upload.html', error='No file selected')

        if file and file.filename.endswith('.xlsx'):
            filename = 'Mivalt Parts List - Master List (rev 7.7.22).xlsx'
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            return render_template('upload.html', success='File uploaded successfully!')
        else:
            return render_template('upload.html', error='Invalid file type. Please upload an Excel file (.xlsx)')

    return render_template('upload.html')

@app.route('/get_product_data')
def get_product_data():
    try:
        screw_press_handler = ScrewPressHandler()
        data = screw_press_handler.read_sheet_data()
        return jsonify({"success": True, "data": data})
    except FileNotFoundError as e:
        app.logger.error(f"Excel file not found: {str(e)}")
        return jsonify({
            "success": False, 
            "error": str(e)
        }), 404
    except Exception as e:
        app.logger.error(f"Error processing Excel data: {str(e)}")
        return jsonify({
            "success": False,
            "error": f"Error processing Excel data: {str(e)}"
        }), 500

@app.route('/preview', methods=['POST'])
def preview():
    form_data = request.form.to_dict()

    # Handle logo upload if present
    if 'company-logo' in request.files:
        logo = request.files['company-logo']
        if logo and logo.filename:
            filename = secure_filename(logo.filename)
            logo_path = os.path.join(LOGO_FOLDER, filename)
            logo.save(logo_path)
            form_data['company-logo'] = f'/images/logos/{filename}'

    return render_template('preview.html', data=form_data)

@app.route('/images/logos/<filename>')
def serve_logo(filename):
    return send_from_directory(LOGO_FOLDER, filename)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)