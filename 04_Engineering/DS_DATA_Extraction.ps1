
$files = Get-ChildItem -Path "W:\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\DS_Templates\EPCIC-12" -Filter item.xml -Recurse

"file_name;data_source;Attribute" | Out-File -FilePath  "W:\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\DS_Templates\EPCIC12_Property_req.csv"
$data = @()
$i = 0
foreach ($file in $files) {
    $i++
    $file.fullname
    Write-Progress -Activity $file.name -PercentComplete ($i / $files.Count * 100)
    [xml]$XmlDocument = Get-Content $file.fullname
    $XmlNamespace = @{ 
        a="http://www.aveva.com/xml/DatasheetMapping"
        b="http://schemas.datacontract.org/2004/07/AVEVA.Datasheets.UI"}
    
    $xpath = "/a:WorkbookConfiguration/a:SheetConfigurations/b:SheetConfiguration/b:CellMapping/b:DatasheetTemplateBaseCell"
    $nodes = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
    
    foreach ($node in $nodes.Node) {
        if ([string]::IsNullOrEmpty($node.Attribute)) {
            CONTINUE
        }
        $data += $file.FullName.Split('\')[9] + ";" + $node.DataSource + ";" + $node.Attribute
    }

}
$data | Select-Object -Unique | Out-File -FilePath  "W:\Appli\DigitalAsset\MP\RUYA_data\Source\AEng\DS_Templates\EPCIC12_Property_req.csv" -Append



