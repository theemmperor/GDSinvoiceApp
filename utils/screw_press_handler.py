import pandas as pd
import os
from pathlib import Path

class ScrewPressHandler:
    def __init__(self):
        # Use Path for better cross-platform compatibility
        self.file_path = Path(os.path.expanduser("~/Spencer Folder/Mivalt Parts List - Master List (rev 7.7.22).xlsx"))
        self.sheet_name = "Screw Presses"

    def read_sheet_data(self):
        """
        Reads Screw Presses sheet from Excel file and returns data as dictionary

        Excel Structure:
        Sheet: Screw Presses
        Columns:
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
        """
        try:
            if not self.file_path.exists():
                # Try alternative path
                alt_path = Path("/Users/amysheehan/Spencer Folder/Mivalt Parts List - Master List (rev 7.7.22).xlsx")
                if alt_path.exists():
                    self.file_path = alt_path
                else:
                    raise FileNotFoundError(
                        f"Excel file not found at either:\n"
                        f"1. {self.file_path}\n"
                        f"2. {alt_path}\n"
                        "Please verify the file path and permissions."
                    )

            # Read specific columns from the Excel sheet
            df = pd.read_excel(
                self.file_path,
                sheet_name=self.sheet_name,
                usecols=[
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
                ]
            )

            # Verify the required columns exist
            required_columns = {
                "Item Name (MD 300 Series)",
                "Manufacturer",
                "Mivalt Part Number",
                "Cost USD"
            }
            missing_columns = required_columns - set(df.columns)
            if missing_columns:
                raise ValueError(f"Missing required columns in Excel sheet: {', '.join(missing_columns)}")

            # Convert DataFrame to list of dictionaries
            return {
                'screw_presses': df.to_dict('records')
            }

        except pd.errors.EmptyDataError:
            raise Exception("The Excel file is empty or contains no valid data.")
        except Exception as e:
            if "Sheet 'Screw Presses' not found" in str(e):
                excel_file = pd.ExcelFile(self.file_path)
                raise Exception(f"Sheet 'Screw Presses' not found in the Excel file. Available sheets: {excel_file.sheet_names}")
            raise Exception(f"Error reading Excel file: {str(e)}")