# param (
#     [Parameter(Mandatory=$true)]
#     [String[]]$mdb_list
# )

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "E3D"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"


Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

# Cleanup of files to avoid 'End of File' error, coming from E3D
#Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged\*.csv'
Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\*.log' -Recurse

# $mdb_list = @("RYA-BJ-DC00_MDB_DE","RYA-BH-DC00_MDB_DE","RUYA-EPCIC11")
$mdb_list = @("RYA-BH-DC00_MDB_DE")

    #Forcing custom_evars.bat

Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

####    Extraction of all tags  ######

foreach ($mdb in $mdb_list) {
    $pml_script_path = [System.IO.Path]::Combine($PSScriptRoot,'macro', 'E3D_Export_All_Tags.mac')
    Write-Log -Level INFO -Message "Extracting All Tag Data from $mdb"
    $argumentList = 'PROD E3D init "C:\Program Files (x86)\AVEVA\Everything3D3.1\launch.init" TTY RYA COMPANY/CPYRYA /' + $mdb + ' $m ' + $pml_script_path
    Start-Job -Name $mdb -ScriptBlock {
        
        Start-Process -Filepath "C:\Program Files (x86)\AVEVA\Everything3D3.1\mon.exe" -ArgumentList $args -Wait
   
    } -ArgumentList $argumentList

}

Wait-Job  -Name $mdb_list
Remove-Job  -Name $mdb_list



$finished = $true

Write-Log -Level INFO -Message "Tag Data Extraction is finished" -finished $finished

