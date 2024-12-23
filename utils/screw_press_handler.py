import pandas as pd
import os
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class ScrewPressHandler:
    def __init__(self):
        self.file_path = Path("data/Mivalt Parts List - Master List (rev 7.7.22).xlsx")
        self.sheet_name = "Screw Presses"  # Expected sheet name
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

    def read_sheet_data(self):
        """Reads Screw Presses sheet from Excel file and returns data as dictionary"""
        logger.info(f"Attempting to read Excel file from: {self.file_path}")

        # Create data directory if it doesn't exist
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

        if not self.file_path.exists():
            logger.error(f"Excel file not found at: {self.file_path}")
            raise FileNotFoundError(
                f"Excel file not found at: {self.file_path}\n"
                "Please upload the Excel file first."
            )

        try:
            logger.info(f"Reading sheet '{self.sheet_name}' from Excel file")
            excel_file = pd.ExcelFile(self.file_path)

            # Log available sheets
            logger.info(f"Available sheets in Excel file: {excel_file.sheet_names}")

            # Try to find a sheet that matches or contains "Screw Presses" (case-insensitive)
            available_sheets = excel_file.sheet_names
            matching_sheets = [s for s in available_sheets if "screw" in s.lower() and "press" in s.lower()]

            if matching_sheets:
                self.sheet_name = matching_sheets[0]  # Use the first matching sheet
                logger.info(f"Found matching sheet: {self.sheet_name}")
            else:
                logger.warning(f"No sheet containing 'Screw Presses' found. Available sheets: {available_sheets}")

            df = pd.read_excel(
                self.file_path,
                sheet_name=self.sheet_name,
                na_values=['', 'NA', 'N/A'],  # Add more NA values if needed
                keep_default_na=True
            )

            # Log found columns
            logger.info(f"Columns found in Excel sheet: {df.columns.tolist()}")

            # Verify all required columns exist
            missing_columns = self.required_columns - set(df.columns)
            if missing_columns:
                logger.error(f"Missing columns: {missing_columns}")
                raise ValueError(
                    f"Missing required columns in Excel sheet: {', '.join(missing_columns)}"
                )

            # Remove any rows where the name is empty
            df = df.dropna(subset=['Item Name (MD 300 Series)'], how='any')

            # Clean and standardize the data
            df = df.assign(
                Manufacturer=df['Manufacturer'].fillna('Not Specified').astype(str),
                Mivalt_Part_Number=df['Mivalt Part Number'].fillna('-').astype(str),
                GDS_Part_No=df['GDS Part No'].fillna('-').astype(str),
                Power=df['Power'].fillna('Not Specified').astype(str),
                Material=df['Material'].fillna('Not Specified').astype(str),
                Lead_Time=df['Lead Time'].fillna('Not Specified').astype(str)
            )

            # Convert numeric columns with proper error handling
            for col in ['Cost (Euro)', 'Cost USD', 'Customer 100%']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

            # Convert to records with proper formatting
            records = []
            for _, row in df.iterrows():
                record = {
                    'Item Name (MD 300 Series)': str(row['Item Name (MD 300 Series)']),
                    'Manufacturer': str(row['Manufacturer']),
                    'Mivalt Part Number': str(row['Mivalt Part Number']),
                    'GDS Part No': str(row['GDS Part No']),
                    'Power': str(row['Power']),
                    'Material': str(row['Material']),
                    'Lead Time': str(row['Lead Time']),
                    'Cost (Euro)': float(row['Cost (Euro)']),
                    'Cost USD': float(row['Cost USD']),
                    'Customer 100%': float(row['Customer 100%'])
                }
                records.append(record)

            logger.info(f"Successfully processed {len(records)} rows of data")

            # Log first record for debugging
            if records:
                logger.debug(f"Sample record: {records[0]}")
            else:
                logger.warning("No valid records found in the Excel file")

            return {
                'screw_presses': records
            }

        except pd.errors.EmptyDataError:
            logger.error("The Excel file is empty or contains no valid data")
            raise Exception("The Excel file is empty or contains no valid data.")
        except Exception as e:
            if "No sheet named" in str(e):
                excel_file = pd.ExcelFile(self.file_path)
                available_sheets = ', '.join(excel_file.sheet_names)
                logger.error(f"Sheet '{self.sheet_name}' not found. Available sheets: {available_sheets}")
                raise Exception(
                    f"Sheet '{self.sheet_name}' not found in the Excel file.\n"
                    f"Available sheets: {available_sheets}"
                )
            logger.error(f"Error reading Excel file: {str(e)}")
            raise Exception(f"Error reading Excel file: {str(e)}")