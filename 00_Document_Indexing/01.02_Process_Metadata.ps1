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

$meta_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\indexing\EPC"+$epc+"_meta\*"

#########           WAS DONE IN PREVIOUS STEPS      ##############

# $source_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\indexing\EPC"+$epc+"_Source\*"

# Remove-Item -Path $source_path -Recurse


$applicable_doc_types = @("PID",
                        "DTS",
                        "LIS",
                        "UFD",
                        "BLD",
                        "DRW",
                        "PFD",
                        "SLD",
                        "LAY",
                        "PLT",
                        "ESD",
                        "CEM",
                        "DID",
                        "REG",
                        "GEA",
                        "ISO")


$all_files = Get-ChildItem -Path $meta_path -Filter *.xml -Recurse

$document_Object_Ids = @()
$count = $all_files.Count
$i = 0

Write-Log -Level INFO -Message "Count before applicable types: $count"

foreach($temp in $all_files){

    Write-Progress -Activity "Processing in Progress" -Status "$i% Complete: from $count" -PercentComplete $i
    $i = $i+ 100 / $count
    
    # If exactly same document exist in the target folder then skip
    # if (-Not([System.IO.File]::Exists($temp.FullName.Replace("meta","Source").Replace("_null.xml",".pdf"))))
    # {
    #     Write-Host "PDF file not found for $temp"
    #     CONTINUE
    # }

    # $destination_document_folder = $temp.Directory.Parent.FullName.Replace("meta","Source")
    # if ([System.IO.Directory]::Exists($destination_document_folder)){
    #     $destination_document_folder = $destination_document_folder 

    #     Remove-Item -Path $destination_document_folder -Recurse
    # }

    $XmlNamespace = @{ a = "http://www.aveva.com/VNET/eiwm"}
    [xml]$XmlDocument = Get-Content $temp.FullName
    $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_doc_type']/a:Value"
    $DOC_TYPE = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace

    if($applicable_doc_types.Contains($DOC_TYPE.Node.InnerText)){
    
        $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='cmis:objectId']/a:Value"
        $DOC_Object_id = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
        $document_Object_Ids+=$DOC_Object_id.Node.InnerText
        CONTINUE
    }        
    
}

$path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC"+$epc+"_Output.xml"

$nodecount = 0
$xmlsettings = New-Object System.Xml.XmlWriterSettings
$xmlsettings.Indent = $true
$xmlsettings.OmitXmlDeclaration = $true

$xmlWriter = [System.XML.XmlWriter]::Create($path, $xmlsettings)
$xmlWriter.WriteStartElement("ObjectIds") 

foreach ($item in $document_Object_Ids){
        $xmlWriter.WriteElementString("ObjectId",$item)
        $nodecount++
}
$xmlWriter.WriteEndElement()

$xmlWriter.Flush()
$xmlWriter.Close()

Write-Log -Level INFO -Message "Count after applicable types: $nodecount"
$finished = $true
Write-Log -Level INFO -Message "Finished Processing Metadata" -finished $finished




