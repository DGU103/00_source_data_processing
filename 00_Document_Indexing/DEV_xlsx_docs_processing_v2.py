import os
import pandas as pd
import re

# Define the paths
source_directory = r'W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source'
regex_csv_path = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv'
output_directory = r'W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Temp'

# Load the regular expressions from the CSV file
regex_df = pd.read_csv(regex_csv_path, delimiter=';', encoding="UTF16")
regex_patterns = regex_df['Regexp'].tolist()

# Function to check if a value matches any of the regex patterns
def matches_regex(value, patterns):
    for pattern in patterns:
        if re.search(pattern, str(value)):
            return True
    return False

# Function to process Excel files
def process_excel_file(file_path, patterns):
    matches = []
    # Determine the engine based on file extension
    if file_path.endswith('.xlsx'):
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except:
            print(f'[ERROR] Unable to read {file_path}')
            return matches
    elif file_path.endswith('.xls'):
        try:
            df = pd.read_excel(file_path, engine='xlrd')
        except:
            print(f'[ERROR] Unable to read {file_path}')
            return matches

    else:
        return matches

    # Iterate through all cells in the DataFrame
    for col in df.columns:
        for value in df[col]:
            if matches_regex(value, patterns):
                matches.append((file_path, col, value))
    return matches

# Recursively process all Excel files in the source directory
all_matches = []
for root, dirs, files in os.walk(source_directory):
    for file in files:
        if file.endswith('.xls') or file.endswith('.xlsx'):
            file_path = os.path.join(root, file)
            matches = process_excel_file(file_path, regex_patterns)
            all_matches.extend(matches)

# Export the matches to a CSV file
output_df = pd.DataFrame(all_matches, columns=['File Path', 'Column', 'Matched Value'])
output_csv_path = os.path.join(output_directory, 'matched_values.csv')
output_df.to_csv(output_csv_path, index=False)

print(f"Matching values have been exported to {output_csv_path}")

