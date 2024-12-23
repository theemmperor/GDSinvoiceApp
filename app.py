from flask import Flask, render_template, request, jsonify, send_from_directory, redirect, url_for, flash
from utils.screw_press_handler import ScrewPressHandler
import os
from werkzeug.utils import secure_filename
import logging
import sys

# Configure logging to output to both file and console
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('app.log')
    ]
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = 'data'  # Set upload folder in config
app.secret_key = 'your-secret-key-here'  # Required for flash messages

logger = app.logger

# Configure upload folders
UPLOAD_FOLDER = app.config['UPLOAD_FOLDER']
IMAGE_FOLDER = os.path.join(os.getcwd(), 'static', 'images')  # Changed to use relative path
LOGO_FOLDER = os.path.join(IMAGE_FOLDER, 'logos')

# Create required directories if they don't exist
for folder in [UPLOAD_FOLDER, IMAGE_FOLDER, LOGO_FOLDER]:
    os.makedirs(folder, exist_ok=True)
    logger.info(f"Created/verified directory: {folder}")

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
            try:
                file.save(file_path)
                logger.info(f"Excel file saved successfully at: {file_path}")
                return render_template('upload.html', success='File uploaded successfully!')
            except Exception as e:
                logger.error(f"Error saving file: {str(e)}")
                return render_template('upload.html', error=f'Error saving file: {str(e)}')
        else:
            return render_template('upload.html', error='Invalid file type. Please upload an Excel file (.xlsx)')

    return render_template('upload.html')

@app.route('/get_product_data')
def get_product_data():
    try:
        logger.info("Attempting to read product data")
        screw_press_handler = ScrewPressHandler()
        data = screw_press_handler.read_sheet_data()

        # Log the actual data for debugging
        logger.debug(f"Raw data from handler: {data}")
        logger.info(f"Successfully retrieved {len(data.get('screw_presses', []))} screw press records")

        return jsonify({
            "success": True,
            "data": data
        })
    except FileNotFoundError as e:
        logger.error(f"Excel file not found: {str(e)}")
        return jsonify({
            "success": False,
            "error": "Excel file not found. Please upload the file first.",
            "details": str(e)
        }), 404
    except Exception as e:
        logger.error(f"Error processing Excel data: {str(e)}")
        return jsonify({
            "success": False,
            "error": "Error processing Excel data",
            "details": str(e)
        }), 500

@app.route('/preview', methods=['POST'])
def preview():
    try:
        form_data = request.form.to_dict()
        logger.debug(f"Received form data: {form_data}")

        # Handle logo upload if present
        if 'company-logo' in request.files:
            logo = request.files['company-logo']
            if logo and logo.filename:
                filename = secure_filename(logo.filename)
                logo_path = os.path.join(LOGO_FOLDER, filename)
                logo.save(logo_path)
                form_data['company-logo'] = f'/images/logos/{filename}'

        return render_template('preview.html', data=form_data)
    except Exception as e:
        logger.error(f"Error in preview: {str(e)}")
        return jsonify({
            "success": False,
            "error": "Error generating preview",
            "details": str(e)
        }), 500

@app.route('/images/logos/<filename>')
def serve_logo(filename):
    return send_from_directory(LOGO_FOLDER, filename)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    logger.info(f"Starting Flask server on port {port}")
    app.run(host='0.0.0.0', port=port, debug=True)