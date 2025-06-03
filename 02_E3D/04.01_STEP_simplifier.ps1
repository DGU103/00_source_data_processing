# Start-Transcript -Path 'W:\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\STP\Vendor_PKG_3D_Models\import_ps.log'

# $pml_script_path = 'C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\unname_sube.mac'
# $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-BH-DC00_MDB_DE $m ' + $pml_script_path
# Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList -Wait


$STPs = Get-ChildItem -Path 'W:\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\STP\16-04-2025-Vendor_PKG_3D_Models'
foreach ($stp in $STPs) {
    if ($stp.BaseName -match '\s') {
        Rename-Item -Path $stp.FullName -NewName ($stp.FullName -replace ' ','_')
    }
}

$STPs = Get-ChildItem -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\STP\16-04-2025-Vendor_PKG_3D_Models' 

foreach ($stp in $STPs) {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    # $Tag = $stp.BaseName.split('_')[0]
    if ($stp.Extension -match '\.st[e]?p') {
        $extension = 'STEP'
    }
    
    $pml_script_path = 'C:\Users\mch107\Downloads\gitRep\DBS-PS_Aveva\00_source_data_processing\02_E3D\macro\05_STEP_With_SIMPLIFIER.mac'
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-BH-DC00_MDB_DE $m ' + $pml_script_path + ' ' + $extension + ' ' + $stp.FullName
    Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList -Wait
    # $argumentList
    $stopwatch.Stop()
    Write-host "[INFO] " -NoNewline -ForegroundColor DarkYellow
    Write-host $stp.BaseName -NoNewline 
    Write-host ' Time for simplification ' $stopwatch.Elapsed.Minutes 'mins' -ForegroundColor Green
}

