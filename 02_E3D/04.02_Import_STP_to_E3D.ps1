# Start-Transcript -Path 'W:\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\STP\Vendor_PKG_3D_Models\import_ps.log'

$STPs = Get-ChildItem -Path 'C:\Users\mch107\Downloads\Gas_Turbine_Compressor(GTC)' -Filter *.stp
# $STPs = Get-ChildItem -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\STP\Temp' -Filter *.stp


foreach ($stp in $STPs) {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $timeouted = $null
    $Tag = $stp.BaseName.split('_')[0]
    $pml_script_path = "$PSScriptRoot\macro\Import_stp.mac"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY AIM SYSTEM/XXXXXX /RYA-BJ-DC00_MDB_DE $m ' + $pml_script_path + ' ' + $Tag + ' ' + $stp.FullName
    $proc = Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $argumentList -Wait

    $proc | Wait-Process -Timeout 30 -ErrorAction SilentlyContinue -ErrorVariable timeouted
    if ($timeouted)
    {
        # terminate the process
        $proc | Stop-Process

        Write-host '[ERROR] ' -NoNewline -ForegroundColor Red
        Write-host $stp.BaseName:  ($stp.Length/1048576) ' MB' -NoNewline 
    
    }
    # elseif ($proc.ExitCode -ne 0)
    # {
    #     # update internal error counter
    # }
    # $argumentList
    $stopwatch.Stop()
    Write-host '[INFO] ' -NoNewline -ForegroundColor DarkYellow
    Write-host $stp.BaseName:  ($stp.Length/1048576) ' MB' -NoNewline 
    Write-host ' Time to import ' $stopwatch.Elapsed.Minutes 'mins' -ForegroundColor Green
}
