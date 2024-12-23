import pandas as pd
import os
from pathlib import Path

class ScrewPressHandler:
    def __init__(self):
        self.file_path = Path("data/Mivalt Parts List - Master List (rev 7.7.22).xlsx")
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
            "Customer 100%",
            "Customer 100% 2"  # Added second Customer 100% column
        }

    def read_sheet_data(self):
        """Reads Screw Presses sheet from Excel file and returns data as dictionary"""
        # Create data directory if it doesn't exist
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

        if not self.file_path.exists():
            sample_data = (
                "Example data format:\n"
                "| Item Name (MD 300 Series) | Manufacturer | Mivalt Part Number | GDS Part No | Power | Material | Lead Time | Cost (Euro) | Cost USD | Customer 100% | Customer 100% 2 |\n"
                "|---------------------------|--------------|-------------------|-------------|-------|----------|-----------|-------------|-----------|---------------|----------------|\n"
                "| Screw Press Model A       | ACME Corp    | MVT-001           | GDS-001     | 500W  | Steel    | 2 weeks   | 1000        | 1200     | 2000          | 2500           |"
            )
            raise FileNotFoundError(
                f"Excel file not found at: {self.file_path}\n"
                "Please place the Excel file in the 'data' folder with the following structure:\n"
                f"Required columns: {', '.join(sorted(self.required_columns))}\n"
                f"Sheet name: {self.sheet_name}\n\n"
                f"{sample_data}"
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