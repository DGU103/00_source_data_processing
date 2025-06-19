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

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "E3D"
$finished = $false

$exportdir = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged"

#Forcing custom_evars.bat

    Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

    . "$PSScriptRoot\..\Common_Functions.ps1"

# Cleanup of previous files to start fresh

Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged\*parents.csv'
Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\*.log' -Recurse

$repeat = $true

while ($repeat) {

$refs = @()
$jobs = @()

foreach ($mdb in $mdb_list) {

    $refs += ($mdb | ForEach-Object {($_ -split '-')[1] + '-' + (($_ -split '-')[2] -split '_')[0]})
    
    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot,'macro', '01_E3D_Export_3D_refs_v2.mac')
    Write-Log -Level INFO -Message "Extracting All Parent Tag Data from $mdb"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY RYA COMPANY/CPYRYA /' + $mdb + ' $m ' + $pml_script_path
    $job = Start-Job -Name $mdb -ScriptBlock {
    Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $args -wait} -ArgumentList $argumentList
    $jobs += [PSCustomObject]@{Job = $job}

}

#Monitoring Settings
$timeout = 600
$idlelimit = 600
$csvfiles = Get-ChildItem -path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged\*parents.csv'
$starttime = Get-Date
$lastupdate = Get-Date

while ($true) {

    Start-Sleep -Seconds 10

################################################################################
# --- Check if all jobs are done ---
################################################################################

    $incompletejobs = $jobs | Where-Object {$_.Job.State -ne 'Completed'}

    if ($incompletejobs.Count -eq 0) {break}

################################################################################
# --- Timeout Check ---
################################################################################

    if ((Get-Date) -gt $starttime.AddSeconds($timeout)) { 
        $jobs.Job | Foreach-Object { Stop-Job $_}
        Write-Log -Level INFO -Message "Timeout Reached. Jobs stopped."

     break  
}


################################################################################
# --- File update check ---
################################################################################
    $updated = $false
    foreach ($entry in $jobs) {

       foreach ($file in $csvfiles) {

            if (Test-Path $file) {
                if ((Get-Item $file).LastWriteTime -gt $lastupdate) {
                    $lastupdate = Get-Date
                    $updated = $true
                }
            }
       }
    }

    if (-not $updated -and ((Get-Date) -gt $lastupdate.AddSeconds($idlelimit))) {
        $jobs.Job | Foreach-Object {Stop-Job $_}
       Write-Log -Level INFO -Message "Files not updated for $idlelimit seconds. Jobs stopped."
        break
    }

}

#Cleanup: remove all jobs

$jobs.Job | Foreach-Object {

    if ($_.State -ne 'Completed') {Stop-Job $_}
    Remove-Job $_
}

$mdb_list = @()

foreach ($ref in $refs) {

    $part = ($ref -split '-')[0]

    $file = Get-ChildItem -path $exportdir -Filter *_$($part)_*parents.csv

    Write-Log -Level INFO -Message "FOUND: $file"

if (-not $file) {

            $mdb_list += ('RYA-' + $ref + '_MDB_DE')
        }    
}

if ($mdb_list.Count -eq 0) {$repeat = $false}

}

Get-Job | Stop-Job
Get-Job | Remove-Job

$finished = $true

Write-Log -Level INFO -Message "Extracting Of all parent data is finished" -finished $finished



