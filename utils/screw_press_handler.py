import pandas as pd
import os
from pathlib import Path
import logging
import sys

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ScrewPressHandler:
    def __init__(self):
        # First check environment variable for the file path
        env_path = os.getenv('EXCEL_FILE_PATH')
        if env_path:
            self.file_path = Path(env_path)
            logger.info(f"Using Excel file path from environment variable: {env_path}")
        else:
            # Default to data directory in project root
            self.file_path = Path("data/Mivalt Parts List - Master List (rev 7.7.22).xlsx")
            logger.info(f"No EXCEL_FILE_PATH set, using default path: {self.file_path}")

        self.sheet_name = "Screw Presses"
        self.required_columns = {
            "Item Name (MD 300 Series)",
            "Manufacturer",
            "Mivalt Part Number",
            "GDS Part No",
            "Power",
            "Material",
            "Lead Time",
            "Cost (Euro)",
            "Cost USD",
            "Customer 100%"
        }

        # Verify file existence and accessibility
        self._verify_file_access()

    def _verify_file_access(self):
        """Verify file existence and accessibility"""
        logger.info(f"Verifying file access at: {self.file_path}")

        if not self.file_path.parent.exists():
            raise FileNotFoundError(
                f"Directory not found: {self.file_path.parent}\n\n"
                "To configure the Excel file location:\n"
                "1. Set the EXCEL_FILE_PATH environment variable:\n"
                "   export EXCEL_FILE_PATH=/path/to/your/excel/file.xlsx\n"
                "2. Or place the file in the default location:\n"
                "   {os.path.join('data', 'Mivalt Parts List - Master List (rev 7.7.22).xlsx')}"
            )

        if not self.file_path.exists():
            raise FileNotFoundError(
                f"Excel file not found at: {self.file_path}\n\n"
                "To configure the Excel file location:\n"
                "1. Set the EXCEL_FILE_PATH environment variable:\n"
                "   export EXCEL_FILE_PATH=/path/to/your/excel/file.xlsx\n"
                "2. Or place the file in the default location:\n"
                "   {os.path.join('data', 'Mivalt Parts List - Master List (rev 7.7.22).xlsx')}"
            )

        try:
            with open(self.file_path, 'rb') as f:
                f.read(1)
            logger.info("File access verified successfully")
        except Exception as e:
            raise Exception(f"Error accessing file: {str(e)}")

    def read_sheet_data(self):
        """Reads Screw Presses sheet from Excel file and returns data as dictionary"""
        try:
            logger.info(f"Reading Excel file from: {self.file_path}")
            df = pd.read_excel(
                self.file_path,
                sheet_name=self.sheet_name,
                usecols=list(self.required_columns)
            )

            # Verify all required columns exist
            missing_columns = self.required_columns - set(df.columns)
            if missing_columns:
                raise ValueError(
                    f"Missing required columns in Excel sheet:\n"
                    f"{', '.join(missing_columns)}\n\n"
                    f"Please ensure your Excel file contains all required columns."
                )

            return {
                'screw_presses': df.to_dict('records')
            }

        except pd.errors.EmptyDataError:
            logger.error("The Excel file is empty or contains no valid data.")
            raise Exception("The Excel file is empty or contains no valid data.")
        except Exception as e:
            if "Sheet 'Screw Presses' not found" in str(e):
                excel_file = pd.ExcelFile(self.file_path)
                available_sheets = ', '.join(excel_file.sheet_names)
                logger.error(f"Sheet 'Screw Presses' not found. Available sheets: {available_sheets}")
                raise Exception(
                    f"Sheet 'Screw Presses' not found in the Excel file.\n"
                    f"Available sheets: {available_sheets}\n"
                    "Please ensure the sheet name matches exactly."
                )
            logger.error(f"Error reading Excel file: {str(e)}")
            raise Exception(f"Error reading Excel file: {str(e)}")