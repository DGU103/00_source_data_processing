import os
import re
import polars as pl
import pandas as pd
import sys
import warnings
from datetime import datetime

warnings.filterwarnings("ignore")

def _resolve(var_name: str, default: str) -> str:
    cli = next((arg.split('=', 1)[1] for arg in sys.argv[1:]
                if arg.lower().startswith(var_name.lower() + '=')), None)
    return cli or os.environ.get(var_name) or default

def collect_excel_files(folder_path):
    excel_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xls') or file.endswith('.xlsx'):
                if not re.search(r'CRS', file):
                    excel_files.append(os.path.join(root, file))
    return excel_files

def load_regex_patterns(csv_path):
    df = pd.read_csv(csv_path, delimiter=';', encoding="UTF16")
    return list(zip(df['Regexp'], df['Naming_template_ID']))

def extract_characteristic(xml_content, tag_name):
    pattern = rf"<Characteristic>\s*<Name>{re.escape(tag_name)}</Name>\s*<Value>(.*?)</Value>"
    match = re.search(pattern, xml_content, re.DOTALL)
    return match.group(1).strip() if match else ''

def get_file_metadata(file_dir):
    metadata = {
        'doctitle': '',
        'doctype': '',
        'issuance_code': '',
        'DATE': datetime.now().strftime('%m/%d/%Y'),
        'doc_date': '',
        'issue_reason': '',
        'file_full_path': ''
    }

    try:
        for file in os.listdir(file_dir):
            if file.endswith('_null.xml'):
                xml_path = os.path.join(file_dir, file)
                with open(xml_path, 'r', encoding='utf-8') as f:
                    xml_content = f.read()

                metadata['doctitle'] = extract_characteristic(xml_content, 'title')
                metadata['doctype'] = extract_characteristic(xml_content, 'pjc_doc_type')
                metadata['issuance_code'] = extract_characteristic(xml_content, 'pjc_last_return_code')               
                raw_date = extract_characteristic(xml_content, 'pjc_revision_date')
                metadata['doc_date'] = raw_date.split()[0] if raw_date else ''
                metadata['issue_reason'] = extract_characteristic(xml_content, 'pjc_revision_object')
                metadata['file_full_path'] = xml_path
                break
    except Exception as e:
        print(f"Warning: Could not read metadata from folder {file_dir}: {e}")

    return metadata

def process_excel_files(excel_files, regex_patterns, output_csv_path):
    matches = []
    for file_path in excel_files:
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        print(f"Processing file: {file_path}")
        try:
            if file_path.endswith('.xlsx'):
                alldf = pl.read_excel(file_path, sheet_id=0, engine='calamine', raise_if_empty=False)
            else:
                alldf = pl.read_excel(file_path, sheet_id=0, engine='xlrd', raise_if_empty=False)

            metadata = get_file_metadata(os.path.dirname(file_path))

            for pattern, naming_template_id in regex_patterns:
                regex = re.compile(pattern)
                for index in alldf:
                    df = alldf[index]
                    for column in df.columns:
                        for cell in df[column]:
                            if regex.search(str(cell)):
                                matches.append([
                                    str(cell), file_name, metadata['doctitle'],
                                    metadata['doctype'], metadata['issuance_code'],
                                    naming_template_id, metadata['DATE'], metadata['doc_date'],
                                    metadata['issue_reason'], file_path
                                ])
        except Exception as e:
            print(f"Error processing file {file_path}: {e}")

    matches_df = pd.DataFrame(matches, columns=[
        'Tag_number', 'Document_number', 'doctitle',
        'doctype', 'issuance_code', 'ST',
        'DATE', 'doc_date', 'issue_reason', 'file_full_path'
    ]).drop_duplicates()
    matches_df.to_csv(output_csv_path, index=False)
    print(f"Results saved to {output_csv_path}")

# Resolve paths
folder_path = _resolve('FOLDER_PATH', r'\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing')
regex_csv_path = _resolve('REGEX_CONFIG', r'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv')
output_csv_path = _resolve('OUTPUT_CSV_PATH', r'\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\excel_indexing.csv')

if __name__ == '__main__':
    process_excel_files(
        collect_excel_files(folder_path),
        load_regex_patterns(regex_csv_path),
        output_csv_path
    )
