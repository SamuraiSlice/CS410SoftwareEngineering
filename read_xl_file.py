import pandas as pd

# Read XL file function
def read_excel_file(file_path):
    try:
        # Read XL file by getting its file path
        df = pd.read_excel(file_path)
        
        # show 1st few rows of dataframe
        print("First 5 rows of the Excel sheet:")
        print(df.head())
        
        # Opt., ret. entire dataframe
        return df
        
    except Exception as e:
        print(f"Error reading the Excel file: {e}")

# Ex. usage
if __name__ == "__main__":
    file_path = 'Paragraph_Assignment_Template.xlsx'  # Path to Excel file
    read_excel_file(file_path)
