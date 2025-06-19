param (
    [String[]]$mdb_list = @(
    'RYA-BJ-DC00_MDB_DE',
    'RYA-BH-DC00_MDB_DE',
    "RYA-LA-DC03_MDB_DE",
    "RYA-MA-DC04_MDB_DE",
    "RYA-PA-DC05_MDB_DE",
    "RYA-QA-DC06_MDB_DE",
    "RYA-RA-DC09_MDB_DE",
    "RYA-TA-DC44_MDB_DE",
    "RYA-UA-DC01_MDB_DE",
    "RYA-WA-DC02_MDB_DE",
    "RYA-XA-DC28_MDB_DE")
)

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "E3D"
$finished = $false

$exportdir = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged"

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

# Cleanup of previous files to start fresh

Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged\*Tagged.csv'
Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\*.log' -Recurse

  #Forcing custom_evars.bat

Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

####    Extraction of all tags  ######

$repeat = $true

while ($repeat) {

$refs = @()

foreach ($mdb in $mdb_list) {

    $refs += ($mdb | ForEach-Object {($_ -split '-')[1] + '-' + (($_ -split '-')[2] -split '_')[0]})

    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot,'macro', 'E3D_Export_All_Tags.mac')
    Write-Log -Level INFO -Message "Extracting All Tag Data from $mdb"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY RYA COMPANY/CPYRYA /' + $mdb + ' $m ' + $pml_script_path
    Start-Job -Name $mdb -ScriptBlock {
        
        Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $args -Wait
   
    } -ArgumentList $argumentList

}

Wait-Job  -Name $mdb_list
Remove-Job  -Name $mdb_list

$mdb_list = @()

foreach ($ref in $refs) {

    $part = ($ref -split '-')[0]

    $file = Get-ChildItem -path $exportdir -Filter *_$($part)_*.csv

    Write-Log -Level INFO -Message "FOUND: $file"

if (-not $file) {

            $mdb_list += ('RYA-' + $ref + '_MDB_DE')
        }    
}

if ($mdb_list.Count -eq 0) {$repeat = $false}

}

$finished = $true

Write-Log -Level INFO -Message "Tag Data Extraction is finished" -finished $finished

