# Start-Transcript -Path "D:\TestDevOps\01_Diagrams\log.log"
Write-Host "Script started"
$pml_script_path = [System.IO.Path]::Combine($PSScriptRoot, 'Aveva_Diagrams-Tag_Report.pmlmac')

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "DIAGRAMS"
$global:finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

#Forcing custom_evars.bat

Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

try {

$params = 'PROD Diagrams init "C:\Program Files (x86)\AVEVA\Diagrams14.1.4\diagrams.init" TTY RYA COMPANY/CPYRYA /RYA-ALL-SCHE $m ' + $pml_script_path

Start-Process -FilePath "C:\Program Files (x86)\AVEVA\Diagrams14.1.4\mon.exe"  -ArgumentList $params -Wait -ErrorAction Stop
$finished = $true
Write-Log -Level INFO -Message "Tags Extraction from Diagrams Finished Successfully"

}

catch {

    Write-Output "Error running $($scriptname): $($_.Exception.Message)"
    Write-Log -Level ERROR -Message "Error running $($scriptname): $($_.Exception.Message)"
    throw
}

