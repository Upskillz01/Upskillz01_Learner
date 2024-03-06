# %%
import win32com.client as win32

def execute_excel_macros(file_path, macro_names):
    # Create an instance of Excel application
    excel_app = win32.Dispatch('Excel.Application')
    
    # Make Excel visible (optional)
    excel_app.Visible = True
    
    # Open the Excel file
    workbook = excel_app.Workbooks.Open(file_path)
    
    try:
        # Loop through each macro name and execute it
        for macro_name in macro_names:
            excel_app.Application.Run(macro_name)
    except Exception as e:
        print(f"Error occurred while executing macro '{macro_name}': {e}")
    finally:
        # Save and close the workbook
        workbook.Save()
        workbook.Close()
        
        # Quit Excel application
        excel_app.Quit()

# Example usage:
file_path = r'D:\Users\ssrifqah\Documents\Python\CashBOP_Anomaly\Outlier1\6. Steps to update fact sheet STATsmart.xlsm'
macro_names = ['NewFolderNewExcel']
execute_excel_macros(file_path, macro_names)


