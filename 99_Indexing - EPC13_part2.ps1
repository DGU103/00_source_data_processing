    Write-Host ("Starting Indexing. Please wait...") -ForegroundColor Cyan
  
    <# Update doc metadata #>
    Set-Location "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\"
    <# Indexing #>
    Start-Process powershell {00_Document_Indexing\01.04_Doc_indexing_multyTread.ps1 -epc 13}
	
    <# Indexing postprocessing to add discipline #>
    Start-Process powershell {00_Document_Indexing\01.05_Indexing_result_postProcessing.ps1 -epc 13} -wait -NoNewWindow

    ## AIM-A Section ##
    <# Creating Document Register for AIM #> 
    Start-Process powershell {00_Document_Indexing\02.01_Document_Register_for_AIM.ps1 -epc 13} -wait -NoNewWindow
    <# Update metadata file #> 
    Start-Process powershell {00_Document_Indexing\02.02_Publish_Doc_to_Tag.ps1 -epc 13}  -Wait -NoNewWindow
    <# Copy PDFs files #> 
    Start-Process powershell {00_Document_Indexing\02.03_PDF_copy_to_AIM.ps1 -epc 13}  -Wait -NoNewWindow

    Write-Output "Finish"