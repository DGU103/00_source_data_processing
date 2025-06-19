param (
    # [Parameter(Mandatory=$true)]
    # [String]$mdb_list
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
    "RYA-XA-DC28_MDB_DE"
    )
)

$exportdir = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged"

#Forcing custom_evars.bat

    Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

    . "$PSScriptRoot\..\Common_Functions.ps1"

# Cleanup of previous files to start fresh

Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged\*parents.csv'
Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\*.log' -Recurse

foreach ($mdb in $mdb_list) {

    # $refs += ($mdb | ForEach-Object {($_ -split '-')[1] + '-' + (($_ -split '-')[2] -split '_')[0]})
    
    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot,'macro', '01_E3D_Export_3D_refs_v2.mac')
    Write-Log -Level INFO -Message "Extracting All Parent Tag Data from $mdb"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY RYA COMPANY/CPYRYA /' + $mdb + ' $m ' + $pml_script_path
    $job = Start-Job -Name $mdb -ScriptBlock {Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $args -Wait} -ArgumentList $argumentList 

}
Wait-Job  -Name $mdb_list
# Remove-Job  -Name $mdb_list

$timeout = 600
$idlelimit = 120
$csvfiles = Get-ChildItem -path '$exportdir\*parents.csv'
$starttime = Get-Date
$lastupdate = Get-Date

while ($true) {

    Start-Sleep -Seconds 10

    if ($job.State -eq 'Completed') {break}

    if ((Get-Date) -gt $starttime.AddSeconds($timeout)) { Stop-Job $job; break}


    $updated = $false
    foreach ($file in $csvfiles) {

        if (Test-Path $file) {
            if ((Get-Item $file).LastWriteTime -gt $lastupdate) {
                $lastupdate = Get-Date
                $updated = $true

            }            
        }
    }

    if (-not $updated -and ((Get-Date) -gt $lastupdate.AddSeconds($idlelimit))) {
        Stop-Job $job
        break
    }

}


# Wait-Job  -Name $mdb_list
Remove-Job  -Name $mdb_list

# Get-Job | Stop-Job
# Get-Job | Remove-Job

