import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

############################################################################
############################################################################
####
####    Title:  testPolicyAssertion.py
####    Author: Henry Steele, Senion Systems Librarian, Library Technology Services, Tufts University
####
####    Purpose:
####        This program takes the data that was created in the bulk loan rule simulator, in this repo
####        and a spreadsheet you make, with some information from this simulator output, and a set of policies
####        you enter by saying a given set of users, item policies, and locations, that you assert
####        should all result in the same loan or request rule, or polices
####    Input:
####        - file picker --> the output of your run of the bulk loan rule simlulator, with columns contained therein
####        - assertion file
####            A file that contains columns:
####                - User Groups
####                    - a list of user groups separated by semicolons
####                - Locations
####                    - a list of locations separated by semicolons
####                - Item Polices 
####                    - a list of item policies separate by semicolons
####                - TOU (Loan)
####                    - the TOU that all of the loan scenarios identified by the preceding three columns should resolve to
####                - TOU (Request)
####                    - the TOU that all of the loan scenarios identified by the preceding three columns should resolve to
####                - Policies columns -->
####                    - find a row that implements the policies indicated by the conditions above and paste the values of all these columns into all the following columns.  Use the input spreadsheet as a template
####
####    Output:
####        - a file that contains all rows that match the condidtions for the assertions entered above.   And outputs any rows (scenarios that) *do not* resolve to the desired TOU and/or policy specification
####    Method:
####        - 

def highlight_unique_values(file_path, output_path):
    # Load the spreadsheet into a pandas DataFrame
    df = pd.read_excel(file_path, engine='openpyxl')

    # Ensure column G exists
    if len(df.columns) < 7:  # Column G is the 8th column (0-indexed)
        raise ValueError("Column G does not exist in the spreadsheet.")

    df = df.sort_values(by=['TOU (Loan)'])
    # Get unique values in column G
    unique_values = df.iloc[:, 6].dropna().unique()  # Column H (0-indexed)
    color_map = {}

    # Generate unique colors for each unique value in Column H
    for i, value in enumerate(unique_values):
        # Generate color codes in ARGB format (8-character hex string)
        red = (100 + (i * 50) % 256) % 256
        green = (150 + (i * 30) % 256) % 256
        blue = (200 + (i * 70) % 256) % 256
        color_map[value] = f"FF{red:02X}{green:02X}{blue:02X}"

    # Load workbook and active sheet
    workbook = load_workbook(file_path)
    sheet = workbook.active

    # Iterate through column H and apply fill
    for row in range(2, sheet.max_row + 1):  # Skip header (row 1)
        cell = sheet[f'G{row}']
        value = cell.value
        if value in color_map:
            fill = PatternFill(start_color=color_map[value], end_color=color_map[value], fill_type="solid")
            cell.fill = fill

    # Save the updated workbook
    workbook.save(output_path)
    print(f"File saved with highlighted column H: {output_path}")

# Usage
input_file = "Bulk_Checkout_Request_Results.xlsx"  # Path to your input file
output_file = "Bulk_Checkout_Request_Results - Formatted.xlsx"  # Path to save the output file
highlight_unique_values(input_file, output_file)
