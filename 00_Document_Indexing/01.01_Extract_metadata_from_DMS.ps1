param (
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [string] $epc
)

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "INDEXING"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"


Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

Write-Log -Level INFO -Message "Retrieving full set of documents metadata from MANASA..." 


$cmiconfig = "D:\CMISGateway\NOC-RUYA\EPC"+$epc+"_Metadata.xml"
Start-Process -Filepath "C:\Program Files\AVEVA\AVEVA NET Gateways\Gateway For CMIS\AVEVA.NET.Gateways.CMIS.App.exe" `
-ArgumentList "-cp $cmiconfig" -Wait

$finished = $true
Write-Log -Level INFO -Message "Retrieve of data completed" -finished $finished