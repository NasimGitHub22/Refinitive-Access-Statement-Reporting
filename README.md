# Refinitive-Access-Statement-Reporting
This Python script is an Access Statement Reporting automation tool that compares CSV files from two different reporting periods (current quarter and previous quarter) and produces an Excel report showing user counts, changes, and detailed lists of users by exchange and region.

## What Does This Code Do?

### 1. Gets user input
Prompts you to enter the year (e.g., 2023) and quarter (e.g., Q3).
This is used in the report titles.

### 2. Reads the CSV data
Loads two CSV files:

current_quarter_file_path → the most recent data.
previous_quarter_file_path → the previous quarter’s data.

The CSVs are expected to have at least:

user_name → person or account name.
exch → exchange code or identifier.

### 3. Maps users to regions
Uses user_country_mapping to assign each user_name to:

- UK
- Hong Kong
- Other

Splits both current and previous quarter data into three datasets, one for each region.

### 4. Creates an Excel report
Generates a file called Access_Statement_Reporting.xlsx at output_file_path.

Each region gets its own worksheet:

- UK
- Hong Kong
- Other

### 5. Writes "Table 1" (Current Quarter)
For each exchange in the region:

- Exchange
- User Count (number of unique users this quarter)
- Users (comma-separated list of users)
- User Count Change = Current Quarter count − Previous Quarter count
- Users Removed (users who were in the previous quarter but not in the current one)
- Users Added (users who are new this quarter)
- If no change → both Added/Removed are set to "None"
- Adds a blank column for visual separation (no border, no fill)

### 6. Writes "Table 2" (Previous Quarter)
In the same sheet, starting at column 7 (G):

- Exchange
- User Count
- Users (list)

### 7. Repeats for all regions
The process in steps 5–6 is repeated for UK, Hong Kong, and Other.

### 8. Saves the workbook

The generated Excel has three sheets (UK, Hong Kong, Carbon) and two tables in each sheet:
Current Quarter Table (with changes from last quarter)
Previous Quarter Table (for reference)

## How To Use Access Reporting Python Script

### Install the required libraries:

Make sure you have installed the pandas, os, and xlsxwriter libraries. If not, you can install them using pip:

pip install pandas
pip install xlsxwriter

### Prepare your data files:
Make sure you have the data files for the current quarter and the previous quarter in CSV format.
Specify the file paths for the current and previous quarter files by replacing 'insert file path here' with the actual file paths in the code:

python
Copy
Edit
current_quarter_file_path = r'insert file path here'
previous_quarter_file_path = r'insert file path here'

### Update the user-country mapping:

Review the user_country_mapping dictionary and update it according to your user names and corresponding countries.

Replace the example names and countries with the actual user names and countries in the dictionary:

user_country_mapping = {
    'User1': 'Location',
    'User 2': 'Location',
    ...
}

### Specify the output file path:

Define the output file path for the generated Excel file by replacing 'insert output file path here' with the desired file path:

output_file_path = r'insert output file path here'

### Run the code:

Run the code using a Python interpreter.
Enter the year and quarter when prompted:

Enter the year: 2023
Enter the quarter (Q1, Q2, Q3, or Q4): Q3
The code will generate an Excel file with access statement reporting for the specified year and quarter, including information about user counts, user changes, and user details for each country.

The generated Excel file will be saved at the specified output file path.

Make sure to have the correct file paths, update the user-country mapping, and provide the necessary user input when prompted.
