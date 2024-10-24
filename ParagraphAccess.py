import win32com.client
import pandas as pd
import os

def call_word_macro(teamNo,student, paragraphTitle, tag, startDateTime, endDateTime):
    # Start Word application
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False  # Set to False to run in background

    template_file = f"{configfile_path}\\Business Report Template V0.1.docm"

    # Open a document
    doc = word.Documents.Open(template_file)

    # Call the macro and pass the value
    try:
        word.Application.Run("AssignAccessToContentControl", student,paragraphTitle,tag, startDateTime, endDateTime)
        # Save a copy of the document
        filename_prefix="Business Report Team"
        extension = ".docm"
        target_filename = f"{configfile_path}\\teams\\{filename_prefix}_{teamNo}{extension}"
        doc.SaveAs(target_filename)

        print("Assignment completed and copy of the document created successfully.")
    except Exception as e:
        print(f"Error calling macro: {e}")

    doc.Close()
    word.Quit()

def AssignParagraphs():
    # Load the Excel file
    excel_filename = "Paragraph Assignment Template.xlsx"
    excel_filepath = f"{configfile_path}\\{excel_filename}"
    df = pd.read_excel(excel_filepath)

    print("Total rows:", df.shape[0])
    # Iterate through each row
    for index, row in df.iterrows():
        print(f"Index: {index+1}, Assignment of {row.iloc[1]}")
        call_word_macro(row.iloc[0], row.iloc[1], row.iloc[3], row.iloc[4], row.iloc[5], row.iloc[6])

#configfile_path = input("Please enter the configuration files path: ")
configfile_path = "C:\\CCSU\\510\\Oct20\\collab\\config"
AssignParagraphs()
