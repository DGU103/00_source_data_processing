$pml_script_path = [System.IO.Path]::Combine($PSScriptRoot, 'ADiagrams_Doc_List.pmlmac')

$params = 'PROD Diagrams init "C:\Program Files (x86)\AVEVA\Diagrams14.1.4\diagrams.init" TTY RYA COMPANY/CPYRYA /RYA-ALL-SCHE $m ' + $pml_script_path

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "DIAGRAMS"
$global:finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

try {

Start-Process -FilePath "C:\Program Files (x86)\AVEVA\Diagrams14.1.4\mon.exe"  -ArgumentList $params -Wait -ErrorAction Stop
$finished = $true
Write-Log -Level INFO -Message "Script Finished Successfully"

}

catch {

    Write-Output "Error running $($scriptname): $($_.Exception.Message)"
    Write-Log -Level ERROR -Message "Error running $($scriptname): $($_.Exception.Message)"
    throw
}
