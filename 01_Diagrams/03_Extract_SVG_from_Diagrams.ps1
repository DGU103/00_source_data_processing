param(
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [int]$epc
)
# Start-Transcript -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Diag\log\EXPORT_SVG.log' -Append
Get-Date -Format 'dd/MM/yyyy'
<# Import resulting CSV to use them as a selection criteria in IED Gateway launching and following mapping file update #>
$csv_path = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Diag\00_EPCIC' + $epc + '_DIAG_Doc_List.csv'

$inArray = Import-Csv -Path $csv_path | Where-Object {$_.name -match '\/(WHPR1-MDM4|CPPR1-MDM5|RPBR1-LTE1)-AS(BJ|BH|LA|MA|PA|QA|RA)[A-Z]-(01|02|03|04|05|06|07|08|09|10|11|12|14|15|16|17|18|19|20|21|25|26|27|28|29|30|31|33|34|35|36|37|38|39|40|41|42|43|44|45|46|47|50|51|52|53|54|55|56|57|58|59|61|62|63|64|65|66|67|69|71|73|80|93|XA|XB|XC|XD|XE|XF|XG|YY)-[0-9]{6}-[0-9]{4}[A-Z]{0,1}'}



$doclist = $inArray 

foreach ($doc in $doclist) {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $doc = $doc.NAME

    $ArgumentList = "-TokenPath \\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\05_IED_access_tokens\RYA-ALL-SCHE.mac -Proj RYA -SelectedElements $doc -CfgNames RUYA_MP_DIAG -EIWMContext NOC+AIM -EIWMFileSuffix null -GenerateFolderVNETFIle False -GenerateTriggerStartFile False"
    $test_file = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Diag\SVG\' + $doc.replace('/','') + '_page1.svg'
    Write-Host "Processing the document $doc"
    if (Test-Path $test_file) {
        Write-Host "The document $doc already exist in destination folder, skipping..." -ForegroundColor DarkYellow
        CONTINUE
    }
    if (Test-Path $test_file.Replace('.svg', '.ERROR')) {
        Write-Host "The document $doc already exist in destination folder, skipping..." -ForegroundColor DarkYellow
        CONTINUE
    }

    Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Diagrams14.1.4\AVEVA.IED.Gateway.Diagrams.exe" -ArgumentList $ArgumentList  
    $proc =  Get-Process | Where-Object {$_.Name -match 'AVEVA.IED.Gateway.Diagrams'}
    Wait-Process -Id $proc.Id -Timeout 240 -ErrorAction SilentlyContinue
    Stop-Process -Id $proc.Id -ErrorAction SilentlyContinue
    if ((Test-Path $test_file) -eq $false) {
        $error_file = $test_file.Replace('.svg', '.ERROR')
        New-Item -Path $error_file
    }
    $stopwatch.stop()
    Write-Host $stopwatch.Elapsed.totalseconds
    
}

Get-ChildItem -Path 'W:\Appli\DigitalAsset\MP\RUYA_data\Source\Diag\SVG\*.xml' | Remove-Item

# $parts = 1
# [int] $partSize = [Math]::Round($inArray.count / $parts, 0)
# if ($partSize -eq 0) { throw "$parts sub-arrays requested, but the input array has only $($inArray.Count) elements." }
# $extraSize = $inArray.Count - $partSize * $parts
# $offset = 0
# $jobs_list = @()
# foreach ($i in 1..$parts) {
        
#     $temp = $inArray[$offset..($offset + $partSize + [bool] $extraSize - 1)]
    
#     $job_id = "TAGGED_EPC" + $epc + "_Batch" + $i.ToString()

#     Start-Job -Name $job_id -ScriptBlock {
#         $doclist = $args

#         foreach ($doc in $doclist) {
#             $doc = $doc.NAME

#             $ArgumentList = "-TokenPath \\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\05_IED_access_tokens\RYA-ALL-SCHE.mac -Proj RYA -SelectedElements $doc -CfgNames RUYA_MP_DIAG -EIWMContext NOC+AIM -EIWMFileSuffix null -GenerateFolderVNETFIle False -GenerateTriggerStartFile False"
#             # Write-Host $ArgumentList
#             $test_file = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Diag\SVG\' + $doc.replace('/SCG_','') + '_page1.svg'
#             if (Test-Path $test_file) {
#                 CONTINUE
#             }
#             Write-Host "Processing the document $doc"
#            Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Diagrams14.1.4\AVEVA.IED.Gateway.Diagrams.exe" -ArgumentList $ArgumentList  -wait 
#         }
        
#     } -ArgumentList $temp
#         $jobs_list += $job_id 
    
#     $offset += $partSize + [bool] $extraSize
#     if ($extraSize) { --$extraSize }

# }
# Wait-Job  -Name $jobs_list

# foreach ($job in $jobs_list) {
#     $export += Receive-Job -Name $job
# }


# Get-Job | Remove-Job


# Stop-Transcript