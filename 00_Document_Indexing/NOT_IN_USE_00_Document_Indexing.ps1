
    param (
    [Parameter(Mandatory=$true)]
    [ValidateSet('05','06','11','12','13')]
    [String]$epc
    # [Parameter(Mandatory=$true)]
    # [string]$csv,
    # [Parameter(Mandatory=$true)]
    # [String] $batch,
    # [String] $batch_file
)
# $log_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\logs\Batch_" + $batch + "_EPC_" + $epc + ".log"
# Start-Transcript -Path $log_path
$Host.UI.RawUI.WindowTitle = "Document Indexing for Package EPCIC $epc"

# Set-Location W:\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\
# Set-Location ..\
$ErrorActionPreference = "Stop"

$date = Get-Date -Format 'dd/MM/yyyy'




# $Light_Regex = Import-Csv -Delimiter ";" -Path "W:\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv"
# $full_regex = Import-Csv -Delimiter ";" -Path "W:\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\06_Regexp_configs\Full_regex.csv"
Write-Host "[Info] Getting the documents from folder..." -ForegroundColor Cyan

$files_Dir = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC"+ $epc +"_Source\"
$files_Dir = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Test\"

$files = Get-ChildItem -Path $files_Dir -Filter *.pdf -Recurse #| Where-Object {$_.BaseName -match $document_selection_criteria}
# $csv = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\Temp\EPC"+$epc+"_batch"+$batch+".csv"
  
# $files = Import-Csv $batch_file
# Remove-Item -Path $csv
#ECHO "### Filtering documents base on doc type ###"
Write-Host "[Info] Filtering documents base on doc type..." -ForegroundColor Cyan

Write-Host "[Info] Processign documents..." -ForegroundColor Cyan


# $file = [System.IO.FileInfo]::New($file_path.FullName)

$local_path = $PSScriptRoot
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'UglyToad.PdfPig.dll')
Import-Module $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'UglyToad.PdfPig.DocumentLayoutAnalysis.dll')
Import-Module $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'BouncyCastle.Crypto.dll')
Import-Module $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'itextsharp.dll')
Import-Module $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'itextsharp.pdfa.dll')
Import-Module $dll_path

$date = Get-Date -Format 'dd/MM/yyyy'
Write-Host "[Info] Processign documents..." -ForegroundColor Cyan

$export_results = @("docnum;pagenumber;textlength")
foreach($f in $files)
{	
    if ($f.BaseName -notmatch "CPPR1-MDM5-ASBJA-[0-9]{2}-[A-Z]{1,2}[0-9]{4}-0001") {
        CONTINUE
    }
    Write-Host $f.FullName
    $file = [System.IO.FileInfo]::new($f.FullName)
    
    if (-not[System.IO.File]::Exists($file.FullName.replace('.pdf','_null.xml'))) {
        Write-Host "[ERROR]" -ForegroundColor Red -NoNewline
        Write-Host "Metadata for the $file.FullName is not found. File will be skipped!"
    }
    
    $XmlNamespace = @{ a = "http://www.aveva.com/VNET/eiwm"}
    [xml]$XmlDocument = Get-Content $file.FullName.replace('.pdf','_null.xml')
    $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision_date']/a:Value" 
    $revision_date = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
    $revision_date = $revision_date.Node.'#text'
    $revision_date = $revision_date.Split(" ")[0]
    

    $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_doc_type']/a:Value" 
    $doc_type = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
    if ($doc_type -match "") {
        <# Action to perform if the condition is true #>
    }
    $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision_object']/a:Value" 
    $reaseon_for_issue = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
    if($reaseon_for_issue -match "CLD"){CONTINUE}
    
    
    # try {
        # $PdfReader = New-Object iTextSharp.text.pdf.PdfReader($file.FullName)  
        # $pageCount = $reader.NumberOfPages
        # $pageCount
        # for ($i = 1; $i -le $pageCount; $i++) {
        #         $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($PdfReader, $i)
        #         $text
        #         $len = $text.Length
        #         Write-Host "Page: $i, text len: $len"
        #     }
    # }
    # catch {
    #     Write-Host "[ERROR] " -ForegroundColor Red -NoNewline
    #     Write-Host "Cannot process $file" -ForegroundColor White
    # }
            
    $pdf = [UglyToad.PdfPig.PdfDocument]::Open($file.FullName)
    for( $i = 1; $i -le $pdf.NumberOfPages; $i++){
        
        # try {
            $page = $pdf.GetPage($i)
            $len = $page.Text.Length
            $images = $page.GetImages()
            for ($i = 0; $i -lt $images.Count; $i++) {
                
            }
            # $export_results += $file.BaseName + ';' + $i.ToString() + ';' + $len.ToString()
            # Write-Host "Page: $i, text len: $len"
        #     $page_words = $page.getwords([UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor]::Instance)
        #     $page_words
        # # }
        # catch {
        #     CONTINUE
        # }
    }
            
}


<# Export for Tag reporting #>
if($epc -in @('11','12','13')){$tag_report = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPCIC"+  [string]$epc +"_Vendor_drawing_report.csv"}
elseif($epc -eq '06')     {$tag_report = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\CPP03\Source\Indexing\EPCIC"+  [string]$epc +"_Vendor_drawing_report.csv"}
elseif($epc -eq '05')     {$tag_report = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\EPCIC"+  [string]$epc +"_Vendor_drawing_report.csv"}

$export_results | Out-File -FilePath $tag_report
#  $tags
# if ($batch -ne 0) {
#     $batch_tag_report ="\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Temp\EPCIC" +  [string]$epc +"_indexing_report" +$batch+ ".csv"
#     $tags | Export-Csv -Path $batch_tag_report -NoTypeInformation -Encoding UTF8
# }

# if (Test-Path $tag_report) {
#     $tags | Export-Csv -Path $tag_report -NoTypeInformation -Encoding UTF8 -Append
# }
# else {
#     $tags | Export-Csv -Path $tag_report -NoTypeInformation -Encoding UTF8
# }
# Stop-Transcript