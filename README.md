# Industrial Component Selection and Invoice System

A professional web application for managing screw press component selection and generating detailed invoice proposals.

## Features

- Interactive component selection interface
- Dynamic pricing calculations
- Professional invoice generation
- Excel data processing integration
- Responsive web design
- Comprehensive data visualization

## Technical Stack

- Backend: Flask (Python)
- Frontend: HTML5, CSS3, JavaScript
- Data Processing: Pandas
- File Handling: Excel integration via openpyxl

## Setup Instructions

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Place your Excel data file in the `data` directory:
   - Filename: "Mivalt Parts List - Master List (rev 7.7.22).xlsx"
   - Required sheets: "Screw Presses"

4. Run the application:
   ```bash
   python app.py
   ```

## Directory Structure

```
├── data/                  # Excel data files
├── static/               
│   ├── css/              # Stylesheet files
│   ├── js/               # JavaScript files
│   └── images/           # Image assets
├── templates/            # HTML templates
├── utils/                # Utility modules
└── app.py               # Main application file
```

## Data Requirements

The Excel file should contain the following columns in the "Screw Presses" sheet:
- Item Name (MD 300 Series)
- Manufacturer
- Mivalt Part Number
- GDS Part No
- Power
- Material
- Lead Time
- Cost (Euro)
- Cost USD
- Customer 100%

## License

This project is proprietary and confidential.
