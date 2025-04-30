import pandas as pd  
from openpyxl import load_workbook
from win32com.client import gencache

def load_and_convert_csv(input_csv_path, output_excel_path):
    try:
        # Load CSV
        df = pd.read_csv(input_csv_path, dtype=str)
        
        # Ensure necessary columns are present
        required_columns = ['txtDueDate']  # Add any other necessary columns here
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")
        
        # Convert columns B and M to strings 
        df.iloc[:, 1] = df.iloc[:, 1].astype(str)  # Column B (index 1)
        df.iloc[:, 12] = df.iloc[:, 12].astype(str) # Column M (index 12)

        # Convert 'txtDueDate' to datetime
        df['txtDueDate'] = pd.to_datetime(df['txtDueDate'], format='%d/%m/%Y', errors='coerce', dayfirst=True)

        # Save to Excel
        df.to_excel(output_excel_path, index=False, header=True)

        # Debugging outputs
        print(df['txtDueDate'].dtype)
        print(df.dtypes.value_counts())
        print("CSV converted to Excel successfully.")
        
        return df
        
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def load_excel_workbook(file_path):
    """
    Loads an Excel workbook using openpyxl.
    
    Parameters:
        file_path (str): The path to the Excel file.

    Returns:
        Workbook: An openpyxl Workbook object.
    """
    try:
        wb = load_workbook(file_path)
        print(f"Workbook '{file_path}' loaded successfully.")
        return wb
    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred while loading the workbook: {e}")

def open_excel_with_win32(file_path, visible=False):
    """
    Opens Excel and a single workbook using win32com.
    Automatically handles corrupted gen_py cache.
    
    Args:
        file_path (str): Path to the Excel file.
        visible (bool): Whether to show the Excel app.
    
    Returns:
        tuple: (excel_app, workbook)
    """
    try:
        excel = gencache.EnsureDispatch("Excel.Application")
        excel.Visible = visible

        workbook = excel.Workbooks.Open(file_path)
        print("Excel workbook loaded with win32 lib")        
        return excel, workbook
    
    except Exception as e:
        print(f"[ERROR] Could not open Excel or workbook: {e}")
        raise

def close_excel_with_win32(excel, workbook, save=True):
    try:
        if save:
            workbook.Save()
        workbook.Close(False)
        excel.Quit()
        print("Closed and saved excel workbook")
    except Exception as e:
        print(f"[ERROR] Failed during Excel/workbook cleanup: {e}")
        raise

    

