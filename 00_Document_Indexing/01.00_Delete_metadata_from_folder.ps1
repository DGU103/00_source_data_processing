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

$meta_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\indexing\EPC"+$epc+"_meta"
$source_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\indexing\EPC"+$epc+"_Source"

$inArray_meta = [System.IO.Directory]::GetFiles("$meta_path" , "*", [System.IO.SearchOption]::AllDirectories).Length
$inArray_source = [System.IO.Directory]::GetFiles("$source_path" , "*.pdf", [System.IO.SearchOption]::AllDirectories).Length


if(-not $inArray_meta -or $inArray_meta -eq 0) {
   Write-Log -Level WARN -Message "No Meta found in $meta_path. Exiting."
    return
}

if (-not $inArray_source -or $inArray_source -eq 0 ) {

    Write-Log -Level WARN -Message "No Sources found in $source_path. Exiting."
    return   
}

Write-Log -Level WARN -Message "Deleting $inArray_meta Metadata objects"
Remove-Item -Path "$meta_path\*" -Recurse
Write-Log -Level WARN -Message "Deleting $inArray_source PDF Source objects"
Remove-Item -Path "$source_path\*" -Recurse

$finished = $true
Write-Log -Level INFO -Message "Cleanup of META and SOURCE data completed" -finished $finished


