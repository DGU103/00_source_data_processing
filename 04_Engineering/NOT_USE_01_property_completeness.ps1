
$ism_file = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\RegisterGateway\source data\class library\clib.xml"

[xml]$XmlDocument = Get-Content $ism_file.FullName

$XmlNamespace = @{ a = "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01" 
 nmcltr ="http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Nomenclature/2015/09"}
 
$temp = Import-Csv \\als.local\noc\data\appli\DigitalAsset\MP\RUYA_data\Source\MTR\WeeklyReports\2024-09-29-EPC13_MTR_validation_report_v6.6.csv
$temp = $temp | Where-Object {$_.Validation_Status -eq 'Valid'}

# $props = Import-Csv -Delimiter ";" -Path  "W:\Appli\DigitalAsset\MP\RUYA_data\RegisterGateway\source data\data load\MP\AENG_PROPERTIES\EPCIC13_Property_Register.csv"
# $aeng_props = @{}
# foreach ($prop in $props) {
#     if (-not ($aeng_props[$prop.TagID])) {
#         $aeng_props.Add($prop.TagID, @{})
#     }
#     $aeng_props[$prop.TagID].Add(@{AttributeID = $prop.AttributeID})
# }
class tag_props {
    [string] $id
    [string] $attribute
    [string] $attribute_name
}

$tags = @{}
foreach ($tag in $temp) {
    $tags.Add($tag.Tag_number, $tag.Tag_class)
}
#  break

$count = $tags.Keys.Count
$i = 0
foreach ($tag in $tags.Keys) {
    $export = @()
    Write-Host "$i from $count"
    $i++
    $aeng_class_id = $tags[$tag]
    $xpath = "/a:ClassLibrary/a:Functionals/a:Class/a:Extension[@nmcltr:AENG='"+ $aeng_class_id +"']"
    $TAG_CLASS_XML = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
    foreach ($attribute in $TAG_CLASS_XML.Node.ParentNode.Attributes.Attribute) {
        $xpath = "/a:ClassLibrary/a:Attributes/a:Attribute[@id='"+ $attribute.id +"']"
        $att_xml = Select-Xml -Xml $XmlDocument  -XPath $xpath -Namespace $XmlNamespace

        $item = New-Object tag_props
        $item.id = $tag
        $item.Attribute = $att_xml.Node.Extension.AENG
        $item.attribute_name = $att_xml.Node.Name

        $export += $item
    }
    $export | Export-Csv -Path "W:\Appli\DigitalAsset\MP\RUYA_data\RegisterGateway\source data\data load\MP\AENG_PROPERTIES\EPCIC13_Property_scope.csv"

}
