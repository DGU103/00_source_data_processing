Clear-Host
$ism_file = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\RegisterGateway\source data\class library\" -Filter clib.xml | Sort-Object | Select-Object -Last 1

[xml]$XmlDocument = Get-Content $ism_file.FullName
 $XmlNamespace = @{ a = "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01" 
 nmcltr ="http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Nomenclature/2015/09"
 infointerface = "http://schemas.aveva.com/InformationInterfaces/Extension/Schema/2017/04"}

 $xpath = "/a:ClassLibrary/a:Extension/infointerface:Data_x0020_Sources" #a:Extension[@nmcltr:AENG='"+$tag_class +"']"

 $root = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace

# $base_dir = 'D:\Gatewayprocessing\GDP_StagingArea\'
$base_dir = '\\qamv3-sapp243\GDP\GDP_StagingArea\'
# New-Item -Path $dir -Force

function FunctionName {
    param (
        [System.Xml.XmlNode[]]$nodes
        
    )
    # $dir = $dir + '\'+$node.Name
    $nexnodes = @()
    foreach ($node in $nodes) {
        
        if ($node.Name -eq  "Columns") {
            CONTINUE
        }
        
        $parent = $node
        $dir = ''
        for ($i = 0; $i -lt 5; $i++) {

            if ($parent.Name -match 'Data_x0020_Sources') {BREAK}
            $fisrt = [string]$parent.Name 
            $dir=  $fisrt + '\' + $dir 
            $parent = $parent.ParentNode
        }
        $dir = $base_dir + $dir
        # $dir
        New-Item -Path $dir -Force -ItemType Directory
        $nexnodes += $node.ChildNodes
    }
    if ($nexnodes.Count -eq 0) {BREAK}
   FunctionName($nexnodes)
}

$s = $root.Node.ChildNodes

FunctionName($s)