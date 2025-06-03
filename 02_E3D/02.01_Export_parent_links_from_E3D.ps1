param (
    # [Parameter(Mandatory=$true)]
    # [String]$mdb_list
    [String[]]$mdb_list = @('RYA-BJ-DC00_MDB_DE',
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

    #Forcing custom_evars.bat

    Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

# Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\*.log' -Recurse
Write-Host "Batch is runninng for the following MDBs: $mdb_list"


foreach ($mdb in $mdb_list) {
    
    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot,'macro', '01_E3D_Export_3D_refs_v2.mac')
    Write-Host "Extracting from $mdb"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY RYA COMPANY/CPYRYA /' + $mdb + ' $m ' + $pml_script_path
    Start-Job -Name $mdb -ScriptBlock {Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $args -Wait} -ArgumentList $argumentList 

    # Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList -Wait
}
Wait-Job  -Name $mdb_list
Remove-Job  -Name $mdb_list


# foreach ($mdb in $mdb_list) {
#     $mdb_path = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\' + $mdb + '.log'
#     # $mdb_path
#     if (Test-Path $mdb_path) {
#         Write-Host "Export finished for $mdb"
#         Get-Content -Path $mdb_path | Write-Host
#         Compress-Archive -Path $mdb_path -DestinationPath $mdb_path.Replace('.log','.zip')
#         Remove-Item -Path $mdb_path
#     }
#     else { 
#         Write-Error -Message "Failed to get log for $mdb"
#     }
# }
# 

Get-Job | Stop-Job
Get-Job | Remove-Job

