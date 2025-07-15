param (
    # [Parameter(Mandatory=$true)]
    # [String]$mdb_list
    # [String[]]$mdb_list,
    [String]$projectID
)
# Get-Job | Stop-Job
# Get-Job | Remove-Job

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "E3D"

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

# Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\*.log' -Recurse

$mdb_list = @("RYA-BJ-DC00_MDB_DE","RYA-BH-DC00_MDB_DE","RYA-LA-DC03_MDB_DE","RYA-MA-DC04_MDB_DE","RYA-PA-DC05_MDB_DE","RYA-QA-DC06_MDB_DE","RYA-RA-DC09_MDB_DE","RYA-TA-DC44_MDB_DE","RYA-UA-DC01_MDB_DE","RYA-WA-DC02_MDB_DE","RYA-XA-DC28_MDB_DE")

# $mdb_list = @("RUYA-EPCIC13")
# $mdb_list = @("RUYA-EPCIC12")
# $mdb_list = @("RYA-LA-DC03_MDB_DE","RYA-MA-DC04_MDB_DE","RYA-PA-DC05_MDB_DE","RYA-QA-DC06_MDB_DE","RYA-RA-DC09_MDB_DE") #Campain 1
# $mdb_list = @("RYA-TA-DC44_MDB_DE","RYA-UA-DC01_MDB_DE","RYA-WA-DC02_MDB_DE","RYA-XA-DC28_MDB_DE") #Campain 2
Write-Log -Level INFO -Message "Batch is runninng for the following MDBs: $mdb_list"

if ($projectID -eq 'RYA') {
   $project = 'RYA COMPANY/CPYRYA'
}
elseif ($projectID -eq 'RST') {
    $project = 'RST SYSTEM/XXXXXX'
}
elseif ($projectID -eq 'RAB') {
    $project = 'RAB SYSTEM/XXXXXX'
}

foreach ($mdb in $mdb_list) {

    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot,'macro', '01_NOC-RYA-E3D_3D_Model_Export.mac')
    Write-Log -Level INFO -Message "Extracting from $mdb"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY ' + $project + ' /' + $mdb + ' $m ' + $pml_script_path
    Start-Job -Name $mdb -ScriptBlock {Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $args -Wait} -ArgumentList $argumentList 
    # Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList  
    # Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList -wait 
}
Wait-Job  -Name $mdb_list
Remove-Job  -Name $mdb_list
$date = Get-Date -Format 'dd_MM_yyyy-hh-mm'
foreach ($mdb in $mdb_list) {
    $mdb_path = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\' +$date + $mdb + '.log'
    # $mdb_path
    if (Test-Path $mdb_path) {
        Write-Log -Level INFO -Message "Export finished for $mdb"
        Get-Content -Path $mdb_path | Write-Host
        Compress-Archive -Path $mdb_path -DestinationPath $mdb_path.Replace('.log','.zip')
        Remove-Item -Path $mdb_path
    }
    else { 
        Write-Log -Level ERROR -Message "Failed to get log for $mdb"
    }
}


# Get-Job | Stop-Job
# Get-Job |  Remove-Job

