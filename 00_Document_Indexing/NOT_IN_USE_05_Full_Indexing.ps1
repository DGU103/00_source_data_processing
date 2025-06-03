    $ErrorActionPreference = "Stop"
    Write-Host ("Retrieving documents from MANASA. Please wait...") -ForegroundColor Cyan
    ## Clear metadat folder->Get all documents metadata from MANASA->
    ## Select required and create Output.xml file->Pull documents renditions from MANASA
    
    Start-Process powershell {.\00_Document_Indexing\01_Update_Documents_in_Source_Folder.ps1} -Wait

    $date = Get-Date -Format "yyyy-MM-dd"
    $source_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing"
    $zip_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Archive\"+ $date + ".zip"
    if (Get-ChildItem -Path $source_path)
    {
        Compress-Archive -Path "$source_path\*.csv" -DestinationPath $zip_path -Force
    }
    Remove-Item -Path "$source_path\*.csv"

    Start-Process powershell {.\00_Document_Indexing\00_Document_Indexing.ps1 -EPCIC_Number 13}
    Start-Process powershell {.\00_Document_Indexing\00_Document_Indexing.ps1 -EPCIC_Number 12} 
    Start-Process powershell {.\00_Document_Indexing\00_Document_Indexing.ps1 -EPCIC_Number 11} -Wait

    ## AIM-A Section ##
    <# Update metadata file #> 
    # Start-Process powershell {00_Document_Indexing/02_doc_metadata_register_for_AIM.ps1} -NoNewWindow -Wait
    <# Copy PDFs files #> 
    # Start-Process powershell {00_Document_Indexing/03_PDF_copy_to_AIM.ps1} -NoNewWindow -Wait

    Write-Output "Finish"