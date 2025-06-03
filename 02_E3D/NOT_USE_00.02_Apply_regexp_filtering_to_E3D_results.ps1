param (
    # [Parameter(Mandatory=$true)]
    [String[]]$epc
)
$epc = @("11","12","13")
$epcs = $epc
# Remove-Item -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\*.log' -Recurse
# Write-Host "Batch is runninng for the following MDBs: $mdb_list"
# $mdb_list = @("RUYA-EPCIC13")
# $mdb_list = @("RUYA-EPCIC13","RUYA-EPCIC12","RUYA-EPCIC11")
$job_lis = @()
foreach ($epc in $epcs) {
    $job_name = "Job_for_" + $epc
    $job_lis += $job_name
    Start-Job -Name $job_name -ScriptBlock {Start-Process powershell "01_E3D_Tagged_Item_full_regex_Regex_Filtering.ps1"  -ArgumentList $args -wait} -ArgumentList $epc
}
Wait-Job  -Name $job_lis
Remove-Job  -Name $job_lis
# foreach ($mdb in $job_lis) {
#     $mdb_path = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\log\'+$mdb+'.log'
#     $mdb_path
#     if (Test-Path $mdb_path) {
#         Write-Host "Export finished for $mdb"
#         Get-Content -Path $mdb_path | Write-Host
#         Remove-Item -Path $mdb_path
#     }
#     else {
#         Write-Error -Message "Export failed for $mdb"
#     }
# }


