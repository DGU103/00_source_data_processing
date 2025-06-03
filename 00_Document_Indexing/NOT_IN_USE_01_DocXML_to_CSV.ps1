
Import-Module C:\Users\mch107\Downloads\importexcel.7.4.1\ImportExcel.psd1


$objects = @()


# $xml_files = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\" -Filter *.xml  -Recurse #| Select-Object -First 1

# $a = Import-Csv -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\DOC_IDs.txt"
# $Path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Output.xml"


# -- Table headers -- #
#"Document_Number;Object_ID"| Out-File -FilePath $Path 


# Define the path to the XLSX file
$xlsxPath = "C:\Users\mch107\Downloads\RPBR1 LTE1 MDDM_14-Apr-25.xlsx"

# Read the XLSX file
$data = Import-Excel -Path $xlsxPath
# Iterate through the rows
$responsible = @()
$responsible += 'Docno;Leader'
foreach ($row in $data) {
    foreach ($column in $row.PSObject.Properties) {
        if ($column.Value -eq 'L') {
              $responsible += $row.docno +';'+ $column.Name
            
        }
    }
}
$responsible | Out-File -FilePath "C:\Users\mch107\Downloads\EPCIC12_responsible.csv" -Encoding utf8
break

$count = $xml_files.Count
$i = 0

#$watch = New-Object System.Diagnostics.Stopwatch
# $objects += "<ObjectIds>"
foreach($xml in $xml_files){
    #  Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
    #  $i =$i+ 100 / $count
  
    $XmlNamespace = @{ a = "http://www.aveva.com/VNET/eiwm"}

    [xml]$XmlDocument = Get-Content $xml.FullName
    $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_doc_type']/a:Value"
    $DOC_NAME = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
    # $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='cmis:objectId']/a:Value"
    # $DOC_OBJECT = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
    # foreach($record in $a){
    #   if($record -match $DOC_NAME.Node.InnerText){   $objects += "<ObjectId>" + $DOC_OBJECT.Node.InnerText + "</ObjectId>"
    #   break}
    # }
    if ($DOC_NAME.Node.InnerText -match 'DTS') {
      Copy-Item -Path $xml.FullName.Replace('_null.xml', '.pdf') -Destination 'C:\Users\mch107\Downloads\forCopilot'
    }
}
# $objects += "</ObjectIds>"

# $objects | Out-File -FilePath $Path #-Append


#$csv_path = $Path

#$csv = Import-Csv $csv_path -Delimiter ";"
#$xlsx_path = $csv_path.Replace(".csv",".xlsx")
