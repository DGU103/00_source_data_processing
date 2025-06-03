
# $mac_path = [System.IO.Path]::Combine($PSScriptRoot, '05_ProjectReplication.mac')
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
        "RYA-XA-DC28_MDB_DE"
        )
)

Write-Host "Batch is runninng for the following MDBs: $mdb_list"

foreach ($mdb in $mdb_list) {
    
    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot, '05_ProjectReplication.mac')
    Write-Host "Replicating $mdb"
    $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m ' + $pml_script_path +' '+ $mdb 

    Start-Process -Filepath "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait
}



# # Get-Job | Stop-Job
# # Get-Job |  Remove-Job




# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-BJ-DC00_MDB_DE'
# Start-Process -Filepath "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m \\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\07_ADMIN\05_ProjectReplication.mac RYA-BH-DC00_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-LA-DC03_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-MA-DC04_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-PA-DC05_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-QA-DC06_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-RA-DC09_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-TA-DC44_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-UA-DC01_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-WA-DC02_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY AIM SYSTEM/XXXXXX /def $m '+ $PSScriptRoot+'\05_ProjectReplication.mac RYA-XA-DC28_MDB_DE'
# Start-Process  "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList -Wait

# 
# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-BJ-DC00_MDB_DE $m '+ $PSScriptRoot+'\..\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-BH-DC00_MDB_DE $m \\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-MA-DC04_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-PA-DC05_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-QA-DC06_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# #### ---- rerun ---- ####
# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-RA-DC09_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-TA-DC44_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-UA-DC01_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-WA-DC02_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 

# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-XA-DC28_MDB_DE $m C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\3d_Model_cleanup.mac'
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList 






# $argumentList = 'PROD ADMIN init "C:\Program Files\AVEVA\Administration2.1\admin.init" TTY 
#  -project=AIM -user=SYSTEM -pass=XXXXXX -mdb=/def -macro=C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\07_ADMIN\05_ProjectReplication.mac'
# Start-Process -Filepath "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList $argumentList

# Start-Process -Filepath "C:\Program Files\AVEVA\Administration2.1\mon.exe" -ArgumentList "-proj=AIM -username=SYSTEM -pass=XXXXXX -tty"