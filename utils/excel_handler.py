import pandas as pd
import os

class ExcelHandler:
    def read_excel(self, file_path):
        """
        Reads Excel file and returns data as dictionary

        Required Excel structure:
        1. File name: data.xlsx
        2. Sheets required:
           - Products:
             Columns: id (int), name (str), price (float)
             Example:
             | id | name          | price  |
             |----|---------------|--------|
             | 1  | Product A     | 99.99  |

           - Customers:
             Columns: id (int), name (str), address (str)
             Example:
             | id | name        | address           |
             |----|-------------|-------------------|
             | 1  | John Doe    | 123 Main St      |
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(
                f"Excel file not found at {file_path}. Please create the file with the required structure."
            )

        try:
            # Verify both sheets exist
            excel_file = pd.ExcelFile(file_path)
            required_sheets = {'Products', 'Customers'}
            missing_sheets = required_sheets - set(excel_file.sheet_names)

            if missing_sheets:
                raise ValueError(f"Missing required sheets: {', '.join(missing_sheets)}")

            products_df = pd.read_excel(file_path, sheet_name='Products')
            customers_df = pd.read_excel(file_path, sheet_name='Customers')

            # Verify required columns
            product_cols = {'id', 'name', 'price'}
            customer_cols = {'id', 'name', 'address'}

            missing_product_cols = product_cols - set(products_df.columns)
            missing_customer_cols = customer_cols - set(customers_df.columns)

            if missing_product_cols:
                raise ValueError(f"Products sheet missing columns: {', '.join(missing_product_cols)}")
            if missing_customer_cols:
                raise ValueError(f"Customers sheet missing columns: {', '.join(missing_customer_cols)}")

            return {
                'products': products_df.to_dict('records'),
                'customers': customers_df.to_dict('records')
            }
        except Exception as e:
            raise Exception(f"Error reading Excel file: {str(e)}")