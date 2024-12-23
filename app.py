from flask import Flask, render_template, request, jsonify, send_from_directory
from utils.screw_press_handler import ScrewPressHandler
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
screw_press_handler = ScrewPressHandler()

# Configure upload folders for macOS
IMAGE_FOLDER = os.path.expanduser('~/Desktop/invoice_images')
LOGO_FOLDER = os.path.join(IMAGE_FOLDER, 'logos')

# Create required directories if they don't exist
for folder in [IMAGE_FOLDER, LOGO_FOLDER]:
    os.makedirs(folder, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_product_data')
def get_product_data():
    try:
        data = screw_press_handler.read_sheet_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

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