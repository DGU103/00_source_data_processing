# Define the path to the Python executable
$python = "C:\Path\To\Python\python.exe"

# Define the path to your Python script
$script = "W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\00_Document_Indexing\DEV_xlsx_docs_processing_v2.py"

# Run the Python script
# @REM & $python $script
Start-Process cmd.exe -ArgumentList "C:\ProgramData\anaconda3\Scripts\activate.bat"

