from flask import Flask, render_template, request, jsonify, send_from_directory
from utils.excel_handler import ExcelHandler
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
excel_handler = ExcelHandler()

# Configure upload folders
EXCEL_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'excel')
IMAGE_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'images')
LOGO_FOLDER = os.path.join(IMAGE_FOLDER, 'logos')

# Create required directories if they don't exist
for folder in [EXCEL_FOLDER, IMAGE_FOLDER, LOGO_FOLDER]:
    os.makedirs(folder, exist_ok=True)

"""
Expected directory structure:
data/
├── excel/
│   └── data.xlsx  # Excel file with 'Products' and 'Customers' sheets
└── images/
    └── logos/     # Store company logos here
"""

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_excel_data')
def get_excel_data():
    excel_file = os.path.join(EXCEL_FOLDER, 'data.xlsx')
    if not os.path.exists(excel_file):
        return jsonify({
            "success": False,
            "error": f"Excel file not found. Please create '{excel_file}' with 'Products' and 'Customers' sheets"
        })

    try:
        data = excel_handler.read_excel(excel_file)
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