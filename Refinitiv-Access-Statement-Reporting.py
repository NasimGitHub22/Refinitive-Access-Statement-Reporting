import pandas as pd
import os
import xlsxwriter

# Prompt the user for the year and quarter
year = input("Enter the year: ")
quarter = input("Enter the quarter (Q1, Q2, Q3, or Q4): ").upper()

# Define the file paths for the current quarter and previous quarter files
current_quarter_file_path = r'insert file path here'
previous_quarter_file_path = r'insert file path here'

# Read the current quarter file
current_quarter_df = pd.read_csv(current_quarter_file_path)

# Read the previous quarter file
previous_quarter_df = pd.read_csv(previous_quarter_file_path)

# ==============================================
# Define the user_name to location mapping
# Fill this dictionary before running the script
# ==============================================
user_country_mapping = {
    'User 1': 'UK',
    'User 2': 'UK',
    'User 3': 'Other',
    'User 4': 'UK',
    'User 5': 'Hong Kong',
    'User 6': 'UK',
    'User 7': 'Hong Kong',
    'User 8': 'Hong Kong',
    'User 9': 'UK',
    'User 10': 'UK'
}

# Split the data based on the location mapping
uk_data_current_quarter = current_quarter_df[current_quarter_df['user_name'].map(user_country_mapping) == 'UK']
hk_data_current_quarter = current_quarter_df[current_quarter_df['user_name'].map(user_country_mapping) == 'Hong Kong']
other_data_current_quarter = current_quarter_df[current_quarter_df['user_name'].map(user_country_mapping) == 'Other']

uk_data_previous_quarter = previous_quarter_df[previous_quarter_df['user_name'].map(user_country_mapping) == 'UK']
hk_data_previous_quarter = previous_quarter_df[previous_quarter_df['user_name'].map(user_country_mapping) == 'Hong Kong']
other_data_previous_quarter = previous_quarter_df[previous_quarter_df['user_name'].map(user_country_mapping) == 'Other']

# Create a new workbook
output_file_path = r'insert output file path here'
workbook = xlsxwriter.Workbook(output_file_path)

# Define the formats for the headers and data
header_format = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
data_format = workbook.add_format({'border': 1})
empty_column_format = workbook.add_format()

# ===============================
# Function to create each sheet
# ===============================
def create_sheet(sheet, current_data, previous_data, location_name):
    # Table 1 - Current Quarter
    sheet.write(0, 0, f'{location_name} Access Statement Reporting {quarter} {year} (From Excel 1)', header_format)
    sheet.write(1, 0, 'Exchange', header_format)
    sheet.write(1, 1, 'User Count', header_format)
    sheet.write(1, 2, 'Users', header_format)
    sheet.write(1, 3, 'User Count Change', header_format)
    sheet.write(1, 4, 'Users Removed', header_format)
    sheet.write(1, 5, 'Users Added', header_format)
    sheet.write(1, 6, '', empty_column_format)  # Empty column

    for i, exch in enumerate(current_data['exch'].unique(), start=2):
        users = current_data[current_data['exch'] == exch]['user_name'].unique()
        user_count_current_quarter = len(users)

        users_previous_quarter = previous_data[previous_data['exch'] == exch]['user_name'].unique()
        user_count_previous_quarter = len(users_previous_quarter)

        user_count_change = user_count_current_quarter - user_count_previous_quarter
        users_added = list(set(users) - set(users_previous_quarter))
        users_removed = list(set(users_previous_quarter) - set(users))

        sheet.write(i, 0, exch, data_format)
        sheet.write(i, 1, user_count_current_quarter, data_format)
        sheet.write(i, 2, ', '.join(users), data_format)
        sheet.write(i, 3, user_count_change, data_format)

        if user_count_change > 0:
            sheet.write(i, 4, ', '.join(users_removed), data_format)
            sheet.write(i, 5, ', '.join(users_added), data_format)
        elif user_count_change < 0:
            sheet.write(i, 4, ', '.join(users_removed), data_format)
            sheet.write(i, 5, ', '.join(users_added), data_format)
        else:
            sheet.write(i, 4, 'None', data_format)
            sheet.write(i, 5, 'None', data_format)

    # Table 2 - Previous Quarter
    sheet.write(0, 7, f'{location_name} Access Statement Reporting Previous Quarter {year} (From Excel 2)', header_format)
    sheet.write(1, 7, 'Exchange', header_format)
    sheet.write(1, 8, 'User Count', header_format)
    sheet.write(1, 9, 'Users', header_format)

    for i, exch in enumerate(previous_data['exch'].unique(), start=2):
        users = previous_data[previous_data['exch'] == exch]['user_name'].unique()
        sheet.write(i, 7, exch, data_format)
        sheet.write(i, 8, len(users), data_format)
        sheet.write(i, 9, ', '.join(users), data_format)

# Create UK Sheet
create_sheet(workbook.add_worksheet('UK'), uk_data_current_quarter, uk_data_previous_quarter, 'UK')

# Create Hong Kong Sheet
create_sheet(workbook.add_worksheet('Hong Kong'), hk_data_current_quarter, hk_data_previous_quarter, 'Hong Kong')

# Create Other Sheet (previously Carbon)
create_sheet(workbook.add_worksheet('Other'), other_data_current_quarter, other_data_previous_quarter, 'Other')

# Save Workbook
workbook.close()

print('Access Statement Reporting saved successfully.')
