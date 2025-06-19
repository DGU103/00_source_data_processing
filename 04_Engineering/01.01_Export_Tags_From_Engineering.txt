param (
    [Parameter(Mandatory=$true)]
    [ValidateSet('11','12','13')]
    [String]$epc
)

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "ENGINEERING"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

#Forcing custom_evars.bat

Copy-Item "$PSScriptRoot\..\custom_evars.bat" -Destination "C:\Users\Public\Documents\AVEVA\Projects" -Force

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

if ($epc -eq '13') {
    $mdb = "/ADMIN_CPP_(BJ&BK)_DE"
}
elseif ($epc -eq '12') {
    $mdb = "/ADMIN_BH_DE"
}
elseif ($epc -eq '11') {
    $mdb = "/ADMIN_WHP_DE"
}


Write-Host "Script started 01.01_Export_Tags_From_Engineering.ps1"
$pml_script_path = [System.IO.Path]::Combine($PSScriptRoot, 'AENG_Export_All_Tags.mac')

$params = 'PROD Engineering init "C:\Program Files\AVEVA\Engineering15.7.1\engineering.init" TTY RYA COMPANY/CPYRYA ' + $mdb + ' $m ' + $pml_script_path

try {
Start-Process -FilePath "C:\Program Files\AVEVA\Engineering15.7.1\mon.exe"  -ArgumentList $params -Wait -ErrorAction Stop
$finished = $true
Write-Log -Level INFO -Message "Script Finished Successfully" -finished $finished

}

catch {

    Write-Output "Error running $($scriptname): $($_.Exception.Message)"
    Write-Log -Level ERROR -Message "Error running $($scriptname): $($_.Exception.Message)"
    throw
}
