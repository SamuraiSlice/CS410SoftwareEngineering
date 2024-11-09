import pandas as pd

# Function to read an Excel file
def read_excel_file(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Display the first few rows of the dataframe
        print("First 5 rows of the Excel sheet:")
        print(df.head())
        
        # Optionally, return the entire dataframe
        return df
        
    except Exception as e:
        print(f"Error reading the Excel file: {e}")

# Example usage
if __name__ == "__main__":
    file_path = 'Paragraph_Assignment_Template.xlsx'  # Path to Excel file
    read_excel_file(file_path)
