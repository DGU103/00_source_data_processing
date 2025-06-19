    <# ORCHESTATOR SCRIPT. 
    
    We can pass following parameters:
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

    E3D -epc $epc -packingvoke $packingvoke
    Indexing -epc $epc -packingvoke $packingvoke
    Engineering -epc $epc -packingvoke $packingvoke
    Diagrams -epc $epc -packingvoke $packingvoke
    E_I -epc $epc -packingvoke $packingvoke

}

if ($fullrun) {

    E3D -fullrun $fullrun
    Indexing -fullrun $fullrun
    Engineering -fullrun $fullrun
    Diagrams -fullrun $fullrun
    E_I -fullrun $fullrun

 }



