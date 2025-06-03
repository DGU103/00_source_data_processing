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

# Define the list of MDBs and file name prefix based on EPC number
if ($epc -eq 11) {
    $mdb_list = @(
        "RYA-LA-DC03_MDB_DE", "RYA-MA-DC04_MDB_DE", "RYA-PA-DC05_MDB_DE",
        "RYA-QA-DC06_MDB_DE", "RYA-RA-DC09_MDB_DE", "RYA-TA-DC44_MDB_DE",
        "RYA-UA-DC01_MDB_DE", "RYA-WA-DC02_MDB_DE", "RYA-XA-DC28_MDB_DE"
    )
    # EPC11 uses each MDB name as the file name prefix (no single $prop_export_batch prefix)
}
elseif ($epc -eq 12) {
    $mdb_list = @("ADMIN_BH_DE")
    $prop_export_batch = 'EPC12_Property_Register_part_'
}
elseif ($epc -eq 13) {
    $mdb_list = @("ADMIN_CPP_(BJ&BK)_DE")
    $prop_export_batch = 'EPC13_Property_Register_part_'
}

$rootpath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\PROPs\"

Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\log\*.log' -Force -ErrorAction SilentlyContinue

Write-Host "Batch is running for the following MDBs: $mdb_list"

# Determine the next batch number to start with, based on existing CSV files

foreach ($mdb in $mdb_list) {

if (($epc -eq 12) -or ($epc -eq 13)) {

    $existing_batches = Get-ChildItem -Path $rootpath `
                       -Filter "$prop_export_batch*.csv" `
                       | ForEach-Object { [int]([regex]::Match($_.Name, '\d+(?!.*\d)').Value) } `
                       | Measure-Object -Maximum `
                       | Select-Object -ExpandProperty Maximum

                        
        if ($existing_batches) {
             $existing_batches++ 
         }
        
         else {
             $existing_batches = 1
         }   

}

if ($epc -eq 11) {
     
    $existing_batches = Get-ChildItem -Path $rootpath `
                           -Filter "$($mdb)_Property_Register_part_*.csv" `
                           | ForEach-Object { [int]([regex]::Match($_.Name, '\d+(?!.*\d)').Value) } `
                           | Measure-Object -Maximum `
                           | Select-Object -ExpandProperty Maximum
            
         if ($existing_batches) {
            $existing_batches++ 
         }
        
         else {
             $existing_batches = 1
         }   

}

    try {
        
        $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot, 'AENG_Export_All_properties.mac')
        $argumentList = 'PROD Engineering init "C:\Program Files\AVEVA\Engineering15.7.1\engineering.init" TTY RYA COMPANY/CPYRYA /' + $mdb + '$m ' + $pml_script_path + ' ' + $existing_batches
 
        Start-Process -FilePath "C:\Program Files\AVEVA\Engineering15.7.1\mon.exe" -ArgumentList $argumentList -Wait

        if (($epc -eq 12) -or ($epc -eq 13)) {
        $latest_batch = Get-ChildItem -Path $rootpath `
        -Filter "$($prop_export_batch)*.csv" `
        | ForEach-Object { [int]([regex]::Match($_.Name, '\d+(?!.*\d)').Value) } `
        | Measure-Object -Maximum `
        | Select-Object -ExpandProperty Maximum

       Write-Log -Level INFO -Message "The last batch record is $latest_batch"
       Write-Log -Level INFO -Message "Already Existing batch record is $existing_batches"

        }

        else {

            $latest_batch = Get-ChildItem -Path $rootpath `
            -Filter "$($mdb)_Property_Register_part_*.csv" `
            | ForEach-Object { [int]([regex]::Match($_.Name, '\d+(?!.*\d)').Value) } `
            | Measure-Object -Maximum `
            | Select-Object -ExpandProperty Maximum

    Write-Log -Level INFO -Message "The last batch record is $latest_batch"
       Write-Log -Level INFO -Message "Already Existing batch record is $existing_batches"

        }

        if ($existing_batches -eq $latest_batch) {

            Write-Log -Level INFO -Message "The full set of data is populated already."
            continue
        }

        if ($epc -eq 11) {

            $path = "$rootpath\$($mdb)_Property_Register_part_$latest_batch.csv"

        if (Select-String -Path $path -Pattern "END OF DATA;;;;" ) {

             Write-Log -Level INFO -Message "Process finished successfully for $mdb"

                 #Removing The 'END OF DATA' Population inside the csv file
        
                 $lines = Get-Content $path
                 $lines[0..($lines.Count - 2)] | Set-Content $path

            }

             else {Write-Log -Level WARN -Message "Data is not finished populating yet"}   

        }


        else {

            $path = "$rootpath\$($prop_export_batch)$latest_batch.csv"
            
       if (Select-String -Path $path -Pattern "END OF DATA;;;;") {

             Write-Log -Level INFO -Message "Process finished successfully for $mdb"
             
                 #Removing The 'END OF DATA' Population inside the csv file
    
                 $lines = Get-Content $path
                 $lines[0..($lines.Count - 2)] | Set-Content $path
                 
        }   
        else {Write-Log -Level WARN -Message "Data is not finished populating yet"}  
        
    }

}

    catch {
        Write-Log -Level WARN -Message "Something went wrong in $mdb"
        Write-Log -Level ERROR -Message "Error: $($_.Exception.Message)"
        throw
    }
}

$finished = $true
Write-Log -Level INFO -Message "Overall process of exporting properties is finished." -finished $finished

