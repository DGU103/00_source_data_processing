
param (
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [int]$epc
)

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "ENGINEERING"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

# Forcing custom_evars.bat
Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

$rootpath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\PROPs\"

# Define the list of MDBs and file name prefix based on EPC number
if ($epc -eq 11) {
    $mdb_list = @(
        "RYA-LA-DC03_MDB_DE", "RYA-MA-DC04_MDB_DE", "RYA-PA-DC05_MDB_DE",
        "RYA-QA-DC06_MDB_DE", "RYA-RA-DC09_MDB_DE", "RYA-TA-DC44_MDB_DE",
        "RYA-UA-DC01_MDB_DE", "RYA-WA-DC02_MDB_DE", "RYA-XA-DC28_MDB_DE"
    )
Remove-Item -Path "$rootpath\*-DC*.txt" -Force -ErrorAction SilentlyContinue 
}
elseif ($epc -eq 12) {
    $mdb_list = @("ADMIN_BH_DE")
Remove-Item -Path "$rootpath\EPC12*.txt" -Force -ErrorAction SilentlyContinue
}

elseif ($epc -eq 13) {
    $mdb_list = @("ADMIN_CPP_(BJ&BK)_DE")
    Remove-Item -Path "$rootpath\EPC13*.txt" -Force -ErrorAction SilentlyContinue
}

Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\log\*.log' -Force -ErrorAction SilentlyContinue

Write-Host "Batch is running for the following MDBs: $mdb_list"

foreach ($mdb in $mdb_list) {

        $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot, 'AENG_properties_completeness_report.mac')
        Write-Host "Extracting from $mdb"
        $argumentList = 'PROD Engineering init "C:\Program Files\AVEVA\Engineering15.7.1\engineering.init" TTY RYA COMPANY/CPYRYA /' + $mdb + ' $m ' + $pml_script_path
        Start-Job -Name $mdb -ScriptBlock {Start-Process -Filepath "C:\Program Files\AVEVA\Engineering15.7.1\mon.exe" -ArgumentList $args -Wait} -ArgumentList $argumentList 

}

Wait-Job  -Name $mdb_list
Remove-Job  -Name $mdb_list

$finished = $true
Write-Log -Level INFO -Message "Overall process of exporting properties is finished." -finished $finished


    #     if (($epc -eq 12) -or ($epc -eq 13)) {
    #     $latest_batch = Get-ChildItem -Path $rootpath `
    #     -Filter "$($prop_export_batch)*.csv" `
    #     | ForEach-Object { [int]([regex]::Match($_.Name, '\d+(?!.*\d)').Value) } `
    #     | Measure-Object -Maximum `
    #     | Select-Object -ExpandProperty Maximum

    #    Write-Log -Level INFO -Message "The last batch record is $latest_batch"
    #    Write-Log -Level INFO -Message "Already Existing batch record is $existing_batches"

    # }

    #     else {

    #         $latest_batch = Get-ChildItem -Path $rootpath `
    #         -Filter "$($mdb)_Property_report_part_*.csv" `
    #         | ForEach-Object { [int]([regex]::Match($_.Name, '\d+(?!.*\d)').Value) } `
    #         | Measure-Object -Maximum `
    #         | Select-Object -ExpandProperty Maximum

    # Write-Log -Level INFO -Message "The last batch record is $latest_batch"
    #    Write-Log -Level INFO -Message "Already Existing batch record is $existing_batches"

    #     }

    #     if ($existing_batches -eq $latest_batch) {

    #         Write-Log -Level INFO -Message "The full set of data is populated already."
    #         continue
    #     }

    #     if ($epc -eq 11) {

    #         $path = "$rootpath\$($mdb)_Property_report_part_$latest_batch.csv"

    #     if (Select-String -Path $path -Pattern "END OF DATA;;;;" ) {

    #          Write-Log -Level INFO -Message "Process finished successfully for $mdb"

    #              #Removing The 'END OF DATA' Population inside the csv file
        
    #              $lines = Get-Content $path
    #              $lines[0..($lines.Count - 2)] | Set-Content $path

    #         }

    #          else {Write-Log -Level WARN -Message "Data is not finished populating yet"}   

    #     }


    #     else {

    #         $path = "$rootpath\$($prop_export_batch)$latest_batch.csv"
            
    #    if (Select-String -Path $path -Pattern "END OF DATA;;;;") {

    #          Write-Log -Level INFO -Message "Process finished successfully for $mdb"
             
    #              #Removing The 'END OF DATA' Population inside the csv file
    
    #              $lines = Get-Content $path
    #              $lines[0..($lines.Count - 2)] | Set-Content $path
                 
    #     }   
    #     else {Write-Log -Level WARN -Message "Data is not finished populating yet"}  
        
    # }

# }

