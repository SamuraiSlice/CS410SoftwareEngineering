import pandas as pd
import win32com.client

def CreateDocumentForTeam(teamNo):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False  # Set to False to run in background

    template_file = f"{configfile_path}\\Business Report Template V0.1.docm"

    # Open a document
    doc = word.Documents.Open(template_file)
    # Save a copy of the document
    filename_prefix = "Business Report Team"
    extension = ".docm"
    target_filename = f"{configfile_path}\\teams\\{filename_prefix}_{teamNo}{extension}"
    doc.SaveAs(target_filename)
    doc.Close()
    word.Quit()

def TeamFormation():
    # Load the Excel file
    excel_filename = "Team Formation.xlsx"
    excel_filepath = f"{configfile_path}\\{excel_filename}"
    df = pd.read_excel(excel_filepath)

    teams = df.iloc[:, 3].unique()
    print("Total number of teams:", teams.shape[0])
    # Iterate through each row

    filtered_dfs = {}
    # Filter the DataFrame for each unique value
    for teamnum in teams:
        CreateDocumentForTeam(teamnum)
        filtered_dfs[teamnum] = df[df.iloc[:, 3] == teamnum]
        for value, filtered_df in filtered_dfs.items():
            print(f"Filtered DataFrame for value '{value}':")
            print(filtered_df)


configfile_path = "C:\\CCSU\\510\\Oct20\\collab\\config"
TeamFormation()