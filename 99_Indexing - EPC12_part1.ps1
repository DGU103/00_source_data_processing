	Write-Host ("Retrieving documents from MANASA. Please wait...") -ForegroundColor Cyan
    ## Clear metadat folder->Get all documents metadata from MANASA->
    <# Update doc metadata #>
    Set-Location "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\"
	<# Cleanup Metadata #>
	Start-Process powershell {00_Document_Indexing\01.00_Delete_metadata_from_folder.ps1 -epc 12} -wait -NoNewWindow
	<# Extract Metadata #>
    Start-Process powershell {00_Document_Indexing\01.01_Extract_metadata_from_DMS.ps1 -epc 12} -wait -NoNewWindow
    <# Process Metadata #>
    Start-Process powershell {00_Document_Indexing\01.02_Process_Metadata.ps1 -epc 12} -wait -NoNewWindow
    <# Extract PDFs from MANASA #>
    Start-Process powershell {00_Document_Indexing\01.03_Extract_PDFs_from_MANASA.ps1 -epc 12} -wait -NoNewWindow

    Write-Output "Finish"