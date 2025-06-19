import os
import re
import polars as pl
import pandas as pd

import multiprocessing
import time

import warnings
warnings.filterwarnings("ignore")

def collect_excel_files(folder_path):
    excel_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xls') or file.endswith('.xlsx'):
                if not re.search(r'CRS', file):
                    excel_files.append(os.path.join(root, file))
    return excel_files

def load_regex_patterns(csv_path):
    df = pd.read_csv(csv_path,  delimiter=';', encoding="UTF16")
    return df['Regexp'].tolist()

def process_excel_files(excel_files, regex_patterns, output_csv_path):
    matches = []
    for file_path in excel_files:
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        print(f"Processing file: {file_path}")
        try:
            if file_path.endswith('.xlsx'):
                alldf = pl.read_excel(file_path,sheet_id=0, engine='calamine',raise_if_empty=False)
            else:
                alldf = pl.read_excel(file_path,sheet_id=0, engine='xlrd',raise_if_empty=False)
            
            for pattern in regex_patterns:
                regex = re.compile(pattern)
                for index in alldf:
                    df = alldf[index]
                    for column in df.columns:
                        for cell in df[column]:
                            if regex.search(str(cell)):
                                matches.append([str(cell), file_path, file_name])
        except Exception as e:
            print(f"Error processing file {file_path}: {e}")
    
    matches_df = pd.DataFrame(matches, columns=['Match Value', 'File Path', 'File Name'])
    matches_df.to_csv(output_csv_path, index=False)
    print(f"Results saved to {output_csv_path}")

# Define paths
folder_path = r'W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source'
csv_path = r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv'
output_csv_path = r'W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Temp\all_matches.csv'

# Collect Excel files
excel_files = collect_excel_files(folder_path)

# Load regex patterns
regex_patterns = load_regex_patterns(csv_path)

# Process Excel files and save matches
process_excel_files(excel_files, regex_patterns, output_csv_path)

process1 = multiprocessing.Process(target=process_excel_files, args=(excel_files, regex_patterns, output_csv_path))