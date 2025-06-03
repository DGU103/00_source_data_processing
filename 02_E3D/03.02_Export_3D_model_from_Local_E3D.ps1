param (
    [String[]]$mdb_list = @(
        'RYA-BJ-DC00_MDB_DE',
        'RYA-BH-DC00_MDB_DE'#,
        # "RYA-LA-DC03_MDB_DE",
        # "RYA-MA-DC04_MDB_DE",
        # "RYA-PA-DC05_MDB_DE",
        # "RYA-QA-DC06_MDB_DE",
        # "RYA-RA-DC09_MDB_DE",
        # "RYA-TA-DC44_MDB_DE",
        # "RYA-UA-DC01_MDB_DE",
        # "RYA-WA-DC02_MDB_DE",
        # "RYA-XA-DC28_MDB_DE"
        )
)

Write-Host "Batch is runninng for the following MDBs: $mdb_list"

foreach ($mdb in $mdb_list) {
    
    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot,'macro', '02_Export_Local_3D_Model.MAC')
    Write-Host "Extracting from $mdb"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /' + $mdb + ' $m ' + $pml_script_path
    Start-Job -Name $mdb -ScriptBlock {Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $args -Wait} -ArgumentList $argumentList 

    # Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList -Wait
}
Wait-Job  -Name $mdb_list
Remove-Job  -Name $mdb_list


# Get-Job | Stop-Job
# Get-Job |  Remove-Job

