import pandas as pd
import re

def find_header_row(sheet, header_text):
    """
    Finds the row index where the specified header text is located.

    :param sheet: The sheet object from which to find the header row.
    :param header_text: The text to search for in the first row.
    :return: The index of the header row.
    """
    for i, row in sheet.iterrows():
        if row[0] == header_text:
            return i
    return 0

def find_last_row(sheet, column_name):
    """
    Finds the last row index with actual data before any explanatory notes or blank rows.

    :param sheet: The sheet object to search in.
    :param column_name: The name of the column to check for the last row with data.
    :return: The index of the last valid row.
    """
    # Iterate backwards from the end of the DataFrame until a non-empty row is found
    for i in reversed(range(len(sheet))):
        if pd.notnull(sheet.loc[i, column_name]):
            return i
    return None

def clean_spaces(cell):
    """
    Replaces sequences of more than two spaces with a single space in a string.

    :param cell: The string to clean.
    :return: The cleaned string with spaces normalized.
    """
    if isinstance(cell, str):
        return re.sub(r' {2,}', ' ', cell)
    return cell

def remove_all_spaces(cell):
    """
    Removes all spaces from a string.
    """
    if isinstance(cell, str):
        return cell.replace(' ', '')
    return cell

def read_and_clean_excel(file_path, sheet_name, header_text):
    """
    Reads an Excel file, identifies the correct header based on the given text, cleans the spaces,
    and cleans the data.

    :param file_path: The path to the Excel file.
    :param sheet_name: The name of the sheet to extract data from.
    :param header_text: The text to identify the header row.
    :return: A cleaned Pandas DataFrame.
    """
    try:
        # Read the sheet without headers first
        sheet = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # Find the header row
        header_row = find_header_row(sheet, header_text)
        if header_row == 0:
            raise ValueError(f"Header row with '{header_text}' not found")

        # Read the data again with the correct header row
        data = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

        # Find the last valid row
        last_row = find_last_row(data, header_text)
        data = data.loc[:last_row - 4]

    
        # Clean spaces in the data and remove all spaces from 'Security Symbol' column
        for column in data.columns:
            if column == header_text:
                data[column] = data[column].apply(remove_all_spaces)
            else:
                data[column] = data[column].apply(clean_spaces)
        
        data['CG Score'] = data['CG Score'].fillna(0)
        data = data.map(lambda x: ' '.join(x.split()) if isinstance(x, str) else x)

        return data

    except Exception as e:
        print(f"Error: {e}")
        return None


def export_to_csv(data, file_path):
    """
    Exports a DataFrame to a CSV file.

    :param data: Pandas DataFrame to export.
    :param file_path: Path where the CSV file will be saved.
    """
    try:
        data.to_csv(file_path, index=False)
        print(f"Data successfully exported to {file_path}")
    except Exception as e:
        print(f"Error exporting to CSV: {e}")

def export_to_json(data, file_path):
    """
    Exports a DataFrame to a JSON file with UTF-8 encoding.

    :param data: Pandas DataFrame to export.
    :param file_path: Path where the JSON file will be saved.
    """
    try:
        data.to_json(file_path, orient='records', lines=True, force_ascii=False, default_handler=str)
        print(f"Data successfully exported to {file_path}")
    except Exception as e:
        print(f"Error exporting to JSON: {e}")

# Example Usage
file_path = 'PP_001.xlsx'
sheet_name = 'profile'
header_text = 'Security Symbol'

cleaned_data = read_and_clean_excel(file_path, sheet_name, header_text)

if cleaned_data is not None:
    # Export to Excel
    export_to_csv(cleaned_data, 'ListedCompanies.csv')

    # Export to JSON
    export_to_json(cleaned_data, 'ListedCompanies.json')
