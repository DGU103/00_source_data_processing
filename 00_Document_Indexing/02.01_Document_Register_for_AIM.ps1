param (
    [Parameter(Mandatory=$true)]
    [ValidateSet('05','06','11','12','13')]
    [String]$epc
)
class ManasaDocument{
    [String] $pjc_revision
    [String] $pjc_project_code
    [String] $name
    [String] $objectId
    [String] $pjc_discipline
    [String] $title
    [String] $pjc_doc_type
    [String] $pjc_revision_object
    [String] $pjc_originator
    [String] $pjc_project_phase
    [String] $pjc_progress
    [String] $pjc_planned_date
    [String] $pjc_last_return_code
    [String] $pdf_rendition
    [String] $platform
}

if ($epc -eq '13') {
    $package_name = "EPC13"
}
elseif ($epc -eq '12') {
    $package_name = "EPC12"
}
elseif ($epc -eq '11') {
    $package_name = "EPC11"
}

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "AIM"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "Running $scriptname for EPCIC $epc"
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"

if($epc -in @('11','12','13')){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\"}
elseif($epc -eq '06'){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\CPP03\Source\Indexing\"}
elseif($epc -eq '05'){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\"}
 

#$objects +="pjc_revision,pjc_project_code,name,objectId,pjc_discipline,title,pjc_doc_type,pjc_revision_object,pdf_rendition"
$metadata_path = $root_path+$package_name+"_Source"
Write-Host ("Collecting XML files from $metadata_path") -ForegroundColor Cyan

$xml_files = Get-ChildItem -Path $metadata_path -Filter *.xml  -Recurse #| Select-Object -First 1

$count = $xml_files.Count

$objects = New-Object ManasaDocument[] $count
$i = 0

$XmlNamespace = @{ a = "http://www.aveva.com/VNET/eiwm"}
Write-Host ("Processing metadata of EPCIC $epc package") -ForegroundColor Cyan

for ($ii = 0; $ii -lt $count; $ii++)
   { 
    Write-Progress -Activity "Extracting metadata... " -Status "$i% Complete: from $count" -PercentComplete $i
    $i = $i+ 100 / $count

    $doc = New-Object -TypeName ManasaDocument

       [xml]$XmlDocument = Get-Content $xml_files[$ii].FullName
       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='cmis:name']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.name = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='cmis:name']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pdf_rendition = $item.Node.InnerText + '.pdf'

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_revision = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_project_code']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_project_code = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='cmis:objectId']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.objectId = $item.Node.InnerText


       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_discipline']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_discipline = $item.Node.InnerText


       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='title']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.title = $item.Node.InnerText


       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_doc_type']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_doc_type = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_project_phase']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_project_phase = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_originator']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_originator = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_last_return_code']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_last_return_code = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_progress']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_progress = $item.Node.InnerText

       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_planned_date']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_planned_date = $item.Node.InnerText


       $xpath = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision_object']/a:Value"
       $item = Select-Xml -Xml $XmlDocument  -XPath $xpath  -Namespace $XmlNamespace
       $doc.pjc_revision_object = $item.Node.InnerText

       $xfopsite = "//a:Template/a:Object/a:Characteristic[a:Name='fop_site']/a:Value"
       $fopsite = (Select-Xml -Xml $XmlDocument -Xpath $xfopsite -Namespace $XmlNamespace).Node.InnerText
       $xfopsector = "//a:Template/a:Object/a:Characteristic[a:Name='fop_sector']/a:Value"
       $fopsector = (Select-Xml -Xml $XmlDocument -Xpath $xfopsector -Namespace $XmlNamespace).Node.InnerText
       $doc.platform = $fopsite + $fopsector.Substring(0,$fopsector.Length-1) 

       $objects[$ii] = $doc
    
   }

$path = "\\qamv3-sapp243\GDP\GDP_StagingArea\MP\Documents\Metadata\" + $package_name + "_DOC_METADATA.csv"

Write-Log -Level INFO -Message "Exporting Doc Register. Please Wait"
$objects | Export-Csv -Path $path -NoTypeInformation -Force
$finished = $true
Write-Log -Level INFO -Message "Exporting Finished Successfully." -finished $finished
