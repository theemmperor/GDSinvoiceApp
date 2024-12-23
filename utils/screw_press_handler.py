import pandas as pd
import os
from pathlib import Path

class ScrewPressHandler:
    def __init__(self):
        self.possible_paths = [
            Path("data/Mivalt Parts List - Master List (rev 7.7.22).xlsx"),  # Project data directory
            Path(os.path.expanduser("~/Spencer Folder/Mivalt Parts List - Master List (rev 7.7.22).xlsx")),  # Home directory
            Path("/Users/amysheehan/Spencer Folder/Mivalt Parts List - Master List (rev 7.7.22).xlsx")  # Absolute path
        ]
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
        self.file_path = self._find_excel_file()

    def _find_excel_file(self):
        """Try to find the Excel file in possible locations"""
        # Create data directory if it doesn't exist
        os.makedirs("data", exist_ok=True)

        for path in self.possible_paths:
            if path.exists():
                return path

        # If file not found, provide detailed guidance
        sample_structure = (
            "Required Excel structure:\n"
            "File name: Mivalt Parts List - Master List (rev 7.7.22).xlsx\n"
            "Sheet name: Screw Presses\n"
            "Required columns:\n" + 
            "\n".join(f"- {col}" for col in self.required_columns)
        )

        paths_str = "\n".join(f"{i+1}. {str(path)}" for i, path in enumerate(self.possible_paths))
        raise FileNotFoundError(
            f"Excel file not found. Please ensure the file exists in one of these locations:\n"
            f"{paths_str}\n\n"
            f"File Structure Requirements:\n{sample_structure}\n\n"
            "Note: The simplest option is to place the file in the project's 'data' folder."
        )

    def read_sheet_data(self):
        """Reads Screw Presses sheet from Excel file and returns data as dictionary"""
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

            return {
                'screw_presses': df.to_dict('records')
            }

        except pd.errors.EmptyDataError:
            raise Exception("The Excel file is empty or contains no valid data.")
        except Exception as e:
            if "Sheet 'Screw Presses' not found" in str(e):
                excel_file = pd.ExcelFile(self.file_path)
                available_sheets = ', '.join(excel_file.sheet_names)
                raise Exception(
                    f"Sheet 'Screw Presses' not found in the Excel file.\n"
                    f"Available sheets: {available_sheets}\n"
                    "Please ensure the sheet name matches exactly."
                )
            raise Exception(f"Error reading Excel file: {str(e)}")