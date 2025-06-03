    <# ORCHESTATOR SCRIPT. 
    
    We can run it with or without the following parameters:
    -tags : Tags Extraction
    -props : Properties Extraction
    -3d : Everything related to 3d model extraction
    -indexing : Everything related to Indexing
    -epc (11,12 or 13) : Can be executed with the specific EPC

    #>

[CmdletBinding()]
param (

[switch]$tags,
[switch]$props,
[switch]$e3d,
[switch]$indexing,
[String]$epc
)

    $fullrun = $False
    $epc = $False
    $packingvoke = $False

if (!($tags.IsPresent) -and !($props.IsPresent) -and !($e3d.IsPresent) -and !($indexing.IsPresent)) {

    $fullrun = $True
}

elseif ($epc.IsPresent) {$packingvoke = $True}

else {$fullrun = $false}

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
# $script = $MyInvocation.MyCommand.Definition
$global:method = "ORCHESTR"
$finished = $false

#Load Common file with all of the necessary Functions
. "$PSScriptRoot\Common_Functions.ps1"


if ($tags.IsPresent) {

    $tags = $True

    $Eng_tags = Read-Host "Extract Engineering Tags? n OR Specify EPC (11,12,13,all):"

    $Dia_tags = Read-Host "Extract All Diagrams Tags?(y/n):"

    $EI_tags = Read-Host "Exract E&I Tags? n OR Specify EPC (11,12,13,all):"

}

if($props.IsPresent) {

    $props = $True

    $Eng_props = Read-Host "Extract Engineering Props? n OR Specify EPC (11,12,13,all):"

    $EI_props = Read-Host "Extract E&I Props? n OR Specify EPC (11,12,13,all):"
}

if($e3d.IsPresent) {

    $e3d = $True
    $e3d_tags = Read-Host "Extract Tags from E3D?(y/n):"

    $e3d_filters = Read-Host "Filtering tags needed? n OR Specify EPC (11,12,13,all):"

    $e3d_links = Read-Host "Publish 3D links to AIM? n OR Specify EPC (11,12,13,all):"

    $e3d_model = Read-Host "Export 3D Models?(y/n):"

}

if($indexing.IsPresent) {

    $indexing = $True
    $meta_update = Read-Host "Update DOC metadata? n OR Specify EPC (11,12,13,all):"
    $epc_envoke = Read-Host "Which Package to Index? n OR Specify EPC (11,12,13,all):"
    $aim_index = Read-Host "Push Indexed Data to AIM? n OR Specify EPC (11,12,13,all):"
}


if ($e3d) {

    E3D -e3d_tags $e3d_tags -e3d_links $e3d_links -e3d_model $e3d_model -e3d_filters $e3d_filters

}

if ($indexing) {

    Indexing -meta_update $meta_update -epc_envoke $epc_envoke -aim_index $aim_index
}

if ($tags -or $props) {

    Engineering -Eng_tags $Eng_tags -Eng_props $Eng_props
    Diagrams -Dia_tags $Dia_tags
    E_I -EI_tags $EI_tags -EI_props $EI_props

}

if ($packingvoke) {

    Write-Log -Level INFO -Message "Starting Orchestration for EPC $epc" -epc $epc
    Write-Log -Level INFO -Message "Invoking E3D Processing"
    E3D -epc $epc -packingvoke $packingvoke
    Write-Log -Level INFO -Message "Invoking Indexing Processing"
    Indexing -epc $epc -packingvoke $packingvoke
    Write-Log -Level INFO -Message "Invoking Engineering Processing"
    Engineering -epc $epc -packingvoke $packingvoke
    Write-Log -Level INFO -Message "Invoking Diagrams Processing"
    Diagrams -epc $epc -packingvoke $packingvoke
    Write-Log -Level INFO -Message "Invoking E&I Processing"
    E_I -epc $epc -packingvoke $packingvoke
    $finished = $true
    Write-Log -Level INFO -Message "Orchestation for EPC $epc is finished" -finished $finished

}

if ($fullrun) {

    Write-Log -Level INFO -Message "Starting Full Orchestration with $env:USERDOMAIN\$env:USERNAME"
    # Write-Log -Level INFO -Message "Invoking E3D Processing"
    #  E3D -fullrun $fullrun
    Write-Log -Level INFO -Message "Invoking Indexing Processing"
     Indexing -fullrun $fullrun
    Write-Log -Level INFO -Message "Invoking Engineering Processing"
     Engineering -fullrun $fullrun
    Write-Log -Level INFO -Message "Invoking Diagrams Processing" 
     Diagrams -fullrun $fullrun
    Write-Log -Level INFO -Message "Invoking E&I Processing"
     E_I -fullrun $fullrun
     $finished = $true
     Write-Log -Level INFO -Message "Orchestation for all EPCs is finished" -finished $finished 
 }



