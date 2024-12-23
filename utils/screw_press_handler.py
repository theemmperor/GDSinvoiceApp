import pandas as pd
import os
from pathlib import Path

class ScrewPressHandler:
    def __init__(self):
        self.file_path = Path("/Users/amysheehan/Spencer Folder/Mivalt Parts List - Master List (rev 7.7.22).xlsx")
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

    def read_sheet_data(self):
        """Reads Screw Presses sheet from Excel file and returns data as dictionary"""
        if not self.file_path.exists():
            raise FileNotFoundError(
                f"Excel file not found at: {self.file_path}\n"
                "Please ensure the file exists and has the correct permissions."
            )

        try:
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

            # Remove any rows where all required fields are empty
            df = df.dropna(how='all', subset=list(self.required_columns))

            return {
                'screw_presses': df.to_dict('records')
            }

        except pd.errors.EmptyDataError:
            raise Exception("The Excel file is empty or contains no valid data.")
        except Exception as e:
            if "No sheet named" in str(e):
                excel_file = pd.ExcelFile(self.file_path)
                available_sheets = ', '.join(excel_file.sheet_names)
                raise Exception(
                    f"Sheet '{self.sheet_name}' not found in the Excel file.\n"
                    f"Available sheets: {available_sheets}\n"
                    "Please ensure the sheet name matches exactly."
                )
            raise Exception(f"Error reading Excel file: {str(e)}")