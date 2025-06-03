import os
import pandas as pd

import xml.etree.ElementTree as ET
# Example usage

search_folder="W:/Appli/DigitalAsset/MP/RUYA_data/Source/Indexing/EPC13_Source/CPPR1-MDM5-ASBJA-10-H16156-0001"
output_folder="W:/Appli/DigitalAsset/MP/RUYA_data/Source/Indexing/Temp"


header_keywords = ["Equipment No", "EquipmentNo", "Tag No", "TagNo", "Tag Number", "TagNumber", "Line Number"]
total_files = 0
processed_files = 0

# Scan for Excel files
print(f"[INFO] Scanning for Excel files in: {search_folder}")
excel_files = []
for rootfolder, ___, files in os.walk(search_folder):
    for file in files:
        if file.endswith(("_null.xml")):
            
            # Load the XML file
            tree = ET.parse(os.path.join(rootfolder, file))
            root = tree.getroot()

            # Define the namespace
            namespace = {'v1': 'http://www.aveva.com/VNET/eiwm'}

            # Find the 'pjc_doc_type' field and print its value
            pjc_doc_type_element = root.find('.//v1:Characteristic[v1:Name="pjc_doc_type"]/v1:Value', namespace)
            # pjc_doc_type = pjc_doc_type_element.text if pjc_doc_type_element is not None else "Field not found"
            if pjc_doc_type_element.text in ['LST', 'REG', 'LIS']:
                # print(pjc_doc_type_element.text)
                for xlsfile in os.listdir(rootfolder):
                    if xlsfile.endswith((".xls", ".xlsx", ".XLSX", "XLS")) and "CRS" not in xlsfile:
                        excel_files.append(os.path.join(rootfolder, xlsfile))

total_files = len(excel_files)
print(f"[INFO] Found {total_files} Excel file(s) to process.")

if total_files == 0:
    print("[INFO] No files to process. Exiting.")
    

for file_path in excel_files:
    print(f"[INFO] Processing file {processed_files + 1} of {total_files}: {file_path}")
    xls = pd.ExcelFile(file_path, engine='openpyxl' if file_path.endswith('.xlsx') else 'xlrd')

    output_rows = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl' if file_path.endswith('.xlsx') else 'xlrd')

        header_row = None
        id_col = None
        for idx, row in df.iterrows():
            if idx >= 20:
                break
            for col in df.columns[:20]:
                if str(row[col]).strip() in header_keywords:
                    header_row = idx+1
                    id_col = col
                    tag_col_header = str(row[col]).strip()
                    break
            if header_row is not None:
                break

        if header_row is None or id_col is None:
            print(f"[WARN] Header row with ID column not found in sheet '{sheet_name}'. Skipping...")
            continue

        headers = df.iloc[header_row].fillna(f"Column{df.columns}")
        df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl' if file_path.endswith('.xlsx') else 'xlrd',header=header_row)
        dfm = df.dropna(subset=[tag_col_header])
        for idx, row in dfm.iterrows():
            id_value = row[tag_col_header]
            for col in dfm.columns:
                value = row[col]
                if pd.notna(value) and str(value).strip():
                    column_name = headers[col]
                    output_rows.append({
                        'Tag Number': id_value,
                        'Attribute Name': column_name,
                        'Attribute Value': value
                    })

    if output_rows:
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_csv = os.path.join(output_folder, f"{base_name}.csv")
        output_df = pd.DataFrame(output_rows)
        output_df.to_csv(output_csv, index=False, encoding='utf-8')
        print(f"[SUCCESS] Exported to: {output_csv}")

    processed_files += 1



print(f"[INFO] Processed {processed_files} of {total_files} file(s).")



