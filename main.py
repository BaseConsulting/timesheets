import os
import pandas as pd
import unicodedata
import warnings

warnings.filterwarnings("ignore")

#pip install xlsxwriter
#pip install xlrd

# Root folder path
root_folder = "/Users/jaromirbartak/FLO Group s.r.o/FLO_Data_Solutions - Timesheets"
# Writing the final dataframe to an excel file with the sheet name 'Timesheet'
final_file_path = "/Users/jaromirbartak/Consolidated_Timesheets.xlsx"


# Function to get all the excel files from the given folder and subfolders
def get_excel_files(root_folder):
    excel_files = []
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if file == 'timesheet.xlsx' and 'Číselníky' not in file and 'pracovnik' not in file:
                excel_files.append(os.path.join(root, file))
    return excel_files


# Get all the excel files from the folder and subfolders
excel_files = get_excel_files(root_folder)

# List to hold dataframes
all_data_frames_timesheet = []
all_data_frames_vacation = []
# Reading each excel file and appending the data to all_data_frames list
for file in excel_files:
    try:
        print(file)
        # Specify the sheet_name parameter to read only the 'Timesheet' sheet
        df_timesheet = pd.read_excel(file, sheet_name='Timesheet', engine='openpyxl')

        df_vacation = pd.read_excel(file, sheet_name='Dovolená', engine='openpyxl')

        # Filtering the rows based on multiple conditions
        df_timesheet = df_timesheet[(df_timesheet['Rok'] == 2024) &
                (df_timesheet['Hodiny'] != "") &
                (df_timesheet['MD'] != "") &
                (df_timesheet['Hodiny'] > 0) &
                (df_timesheet['MD'] > 0)]

        #solve the czech encoding issues in names
        employee_name = os.path.basename(os.path.dirname(file)).replace("_", " ")
        employee_name_cp1250 = unicodedata.normalize('NFC', employee_name)
        # rearrange first name and lastname
        name_parts = employee_name_cp1250.split(' ')
        employee_name_cp1250 = ' '.join(name_parts[::-1])

        df_timesheet.insert(0, 'Employee', employee_name_cp1250)  # Renaming the first column to 'Employee'
        all_data_frames_timesheet.append(df_timesheet)
        df_vacation.insert(0, 'Employee', employee_name_cp1250)  # Renaming the first column to 'Employee'
        all_data_frames_vacation.append(df_vacation)
    except Exception as e:
        print(f"Could not read file {file} because {e}")

# Concatenating all the dataframes into one
final_df_timesheet = pd.concat(all_data_frames_timesheet, ignore_index=True)
final_df_vacation = pd.concat(all_data_frames_vacation, ignore_index=True)

# order data
sorted_df_timesheet = final_df_timesheet.sort_values(by=["Employee", "Datum"], ascending=[True, True]).reset_index(drop=True)
sorted_df_vacation = final_df_vacation.sort_values(by="Employee", ascending=True).reset_index(drop=True)
#print(sorted_df_vacation)

# Use ExcelWriter to write to multiple sheets
with pd.ExcelWriter(final_file_path, engine='xlsxwriter') as writer:
    # Write each DataFrame to a different sheet
    sorted_df_timesheet.to_excel(writer, sheet_name='Timesheet', index=False)
    sorted_df_vacation.to_excel(writer, sheet_name='Dovolená', index=False)
    # No need to explicitly call save(), it's done when you exit the block

print(f"File created at {final_file_path}")
