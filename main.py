from openpyxl import load_workbook

# File path to my spreadsheet
file_path = "C:/Users/akbba/OneDrive/Desktop/Finance Tracker.xlsx"

try: 
    workbook = load_workbook(file_path)
    # Selecting the active sheet
    worksheet = workbook.active
    print("Spreadsheet opened successfully!")
except FileNotFoundError:
    print("Error: The file was not found. Check the file path.")
except Exception as e:
    print(f"An error occured: {e}")
