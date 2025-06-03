# param(
#     [Parameter(Mandatory=$true)]
#     [ValidateSet(11,12,13)]
#     [int]$epc
# )
$date = Get-Date -Format 'dd/MM/yyyy'
<# Import resulting CSV to use them as a selection criteria in IED Gateway launching and following mapping file update #>
Start-Transcript -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Diag\log\EXPORT_3D_IED_GW '+ $date +'.log' -Append


    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    # $doc = $doc.NAME

    $ArgumentList = "-TokenPath \\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\05_IED_access_tokens\RUYA-EPCIC13.mac -Proj RYA -SelectedElements /BJYY-PVV-DE -CfgNames MP_RUYA_3D_MODEL_EXPORT -EIWMContext NOC+AIM -EIWMFileSuffix null -GenerateFolderVNETFIle False -GenerateTriggerStartFile False"

    Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Diagrams14.1.4\AVEVA.IED.Gateway.Diagrams.exe" -ArgumentList $ArgumentList  
    # $proc =  Get-Process | Where-Object {$_.Name -match 'AVEVA.IED.Gateway.Diagrams'}
    # Wait-Process -Id $proc.Id -Timeout 120 -ErrorAction SilentlyContinue
    # Stop-Process -Id $proc.Id -ErrorAction SilentlyContinue
    $stopwatch.stop()
    Write-Host $stopwatch.Elapsed.totalseconds

