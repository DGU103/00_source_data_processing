param(
    # [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [int]$epc = 13
)
# Set-Location W:\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\
# Set-Location ..\
$ErrorActionPreference = "Stop"

class Tag2Doc{
    [String] $Tag_number
    [String] $short_id
    [String] $Document_number
    [String] $ST
    [String] $DATE
    [string] $doc_date
    [string] $issue_reason
    
}

Import-Module "$PSScriptroot\lib\UglyToad.PdfPig.dll"
Import-Module "$PSScriptroot\lib\UglyToad.PdfPig.DocumentLayoutAnalysis.dll"
Import-Module "$PSScriptroot\lib\BouncyCastle.Crypto.dll"
Import-Module "$PSScriptroot\lib\itextsharp.dll"
Import-Module "$PSScriptroot\lib\itextsharp.pdfa.dll"

# if($epc -eq 11)    {$document_selection_criteria = "WHPR1-MDM4-[A-Z]{5}-[0-9]{2}-[0-9]{6}-[0-9]{4}$"}
# elseif($epc -eq 12){$document_selection_criteria = "RPBR1-LTE1-(ASBHA|ASYYY)-[0-9]{2}-[0-9]{6}-[0-9]{4}$"}
# elseif($epc -eq 13){$document_selection_criteria = "CPPR1-MDM5-[A-Z]{5}-[0-9]{2}-[0-9]{6}-[0-9]{4}$"}

$Host.UI.RawUI.WindowTitle = "Document Indexing for Package EPCIC $epc"

$files_Dir = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC"+ $epc +"_Source\"
$files_Dir = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\EPC05_Source\WHP03-PMC2-ASKAA-10-R25011-0001\0X"

$date = Get-Date -Format 'dd/MM/yyyy'

$Inst_Regexes = @("AAH","AAHH","AI","AIS","AIT","APB","AR","ARC","ASP","AT","BDV","BX","BY","CAM","CMO","CMP","CPF","CPJ","CPR","CS","CVA","CY","DI","DT","EPB","ESDV","EWS","EX","EY","FAH","FAHH","FAL","FALL","FC","FCV","FE","FG","FI","FIT","FIV","FMX","FO","FPS","FQ","FQI","FQV","FQVY","FS","FSH","FSHH","FSL","FSLL","FT","FVI","FX","FY","GD","GDAH","GDAHH","GDR","GDS","GDT","GVA","GVAA","HC","HD","HDAH","HDAHH","HDC","HDR","HDS","HDT","HF","HG","HGAH","HGAHH","HGS","HIT","HR","HRAH","HS","HSS","HT","HVA","IAM","ICD","ID","IMS","IPC","IR","IRAH","JBC","JBE","JBF","JBJ","JBS","LAH","LAHH","LAL","LALL","LC","LCV","LG","LI","LIT","LOS","LRS","LS","LSC","LSD","LSH","LSHH","LSHL","LSL","LSLL","LSS","LT","LVI","LY","MAC","MACA","MCT","MI","MOV","MRD","MT","MWS","OCP","OWS","PA","PAH","PAHH","PAL","PALL","PB","PC","PCD","PCV","PDAH","PDAHH","PDAL","PDALL","PDC","PDCV","PDI","PDIT","PDRC","PDS","PDSH","PDSHH","PDSL","PDSLL","PDT","PDY","PE","PI","PIT","PRI","PRV","PS","PSE","PSH","PSHH","PSL","PSLL","PSV","PT","PV","PVI","PX","PY","R","RCU","RD","RTD","RTU","SAH","SAHH","SAL","SALL","SD","SDAH","SDV","SE","SI","SL","SOV","SS","SSH","SSL","ST","SVC","SVP","SWS","SX","SY","TAH","TAHH","TAL","TALL","TC","TCV","TDAH","TDAL","TDIC","TDY","TE","TES","TI","TIT","TMX","TS","TSH","TSHH","TSHL","TSL","TSLL","TSV","TT","TVI","TW","TY","UA","UV","VAH","VAHH","VDU","VGDAH","VGDAHH","VHDAH","VHDAHH","VHGAH","VHGAHH","VHRAH","VIRAH","VMACA","VSDAH","VT","WAA","WMA","WMH","WML","WMR","WMV","WT","X","XA","XAH","XAHH","XC","XCT","XCV","XEP","XI","XL","XPI","XPS","XS","XT","XY","Y","YSL","ZAH","ZAHH","ZE","ZI","ZIC","ZIO","ZL","ZLC","ZLO","ZS","ZSC","ZSO","ZT", "HCV", "XZSL", "PRVXV") -join "|"


$Light_Regex = Import-Csv -Delimiter ";" -Path "$PSScriptroot\..\06_Regexp_configs\Light_regex.csv"
# $full_regex = Import-Csv -Delimiter ";" -Path "W:\Appli\DigitalAsset\MP\RUYA_data\GitRepo\00_source_data_processing\06_Regexp_configs\Full_regex.csv"
Write-Host "[Info] Getting the documents from folder..." -ForegroundColor Cyan

$files = Get-ChildItem -Path $files_Dir -Filter *.pdf -Recurse #| Where-Object {$_.BaseName -match $document_selection_criteria}

#ECHO "### Filtering documents base on doc type ###"
Write-Host "[Info] Filtering documents base on doc type..." -ForegroundColor Cyan



$tag_report = [IO.Path]::Combine("\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing", ("EPCIC"+  [string]$epc +"_indexing_report.csv"))
$Tag_to_Doc = [IO.Path]::Combine("\\qamv3-sapp243\GDP\GDP_StagingArea\MP\Documents\Tag2Doc", ("EPCIC"+  [string]$epc +"_Tag2Doc.csv"))


Write-Host "[Info] Processign documents..." -ForegroundColor Cyan

$tags = @([Tag2Doc]::new())

$count = $files.count
$index = 0
foreach($file in $files)
{	
    # $File_data = @([Tag2Doc]::new())
    $page_words = @()
    # Write-Progress -Activity "Processing $file. Overall progress:" -Status "$index% Complete: from $count" -PercentComplete $index
    # $index =$index + 100 / $count

    # $XmlNamespace = @{ a = "http://www.aveva.com/VNET/eiwm"}
    # if (-not[System.IO.File]::Exists($file.FullName.replace('.pdf','_null.xml'))) {
    #     Write-Host "[ERROR]" -ForegroundColor Red -NoNewline
    #     Write-Host "Metadata for the $file is not found. File will be skipped!"
    # }
    # [xml]$XmlDocument = Get-Content $file.FullName.replace('.pdf','_null.xml')
    # $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision_date']/a:Value" 
    # $revision_date = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
    # $revision_date = $revision_date.Node.'#text'
    # $revision_date = $revision_date.Split(" ")[0]
    

    # $XPATH = "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision_object']/a:Value" 
    # $reaseon_for_issue = Select-Xml -Xml $XmlDocument  -XPath $XPATH  -Namespace $XmlNamespace
    # $reaseon_for_issue = $reaseon_for_issue.Node.'#text'
    
    try {
        $PdfReader = New-Object iTextSharp.text.pdf.PdfReader($file.FullName)  
        $BookMarks = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($PdfReader)

        foreach($a in $BookMarks){
            foreach($b in $a.Kids ){
                foreach($c in $b.Kids ){
                    foreach($d in $c.Kids ){
                        foreach ($regex in $Light_Regex) {
                            if($d.Title.Split()[0] -match $regex.Regexp){
                                $record = New-Object Tag2Doc
                                $record.Tag_number = $d.Title.Split()[0]
                                $record.document_number = $file.BaseName 
                                $record.ST = "From bookmarks"
                                $record.date = $date
                                $record.doc_date = $revision_date
                                $record.issue_reason = $reaseon_for_issue
                                $tags += $record
                                # $File_data += $record
                                # $tags += $d.Title.Split()[0] +  ";" + $file.BaseName + ";" + "From bookmarks"
                                BREAK
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Host "[ERROR] " -ForegroundColor Red -NoNewline
        Write-Host "Cannot process $file" -ForegroundColor White
    }

    $a = $file -split"-"
    $Tag_number_Prefix = $a[2].Substring(0,4) + "A"

    $pdf = [UglyToad.PdfPig.PdfDocument]::Open($file.FullName)
    for( $i = 1; $i -le $pdf.NumberOfPages; $i++){
        $inst_sequens_numbers = @()
        
        try {
            $page = $pdf.GetPage($i)
            $page_words = $page.getwords([UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor]::Instance)
        }
        catch {
            Write-Host "error" -ForegroundColor Red
            # CONTINUE
        }
        $words_export = $files_Dir+ '\' + $file.BaseName + '.csv'
        $page_words | Export-Csv -Path $words_export
        foreach($word in $page_words){
            
            if([string]::IsNullOrEmpty($word.Text)){CONTINUE}

            foreach ($regex in $Light_Regex) {
                if($word.Text -match $regex.Regexp){
                    $record = New-Object Tag2Doc
                    $record.Tag_number = $word.Text
                    $record.document_number = $file.BaseName 
                    $record.ST = $regex.Naming_template_ID
                    $record.date = $date
                    $record.doc_date = $revision_date
                    $record.issue_reason = $reaseon_for_issue
                    $tags += $record
                    # # $File_data += $record
                    # BREAK
                } 
            }
        }
        if($page.Size -notmatch "A1"){
            CONTINUE
        }
        if ($file.FullName -notmatch '-10-') {
            CONTINUE
        }
        foreach($word in $page_words){
        #                         #Regular Inst Tag number sequence#                       # Temp Inst Tag number sequence#
            if ($word.Text -match "^[0-9]{6}[A-Z]?$" -or $word.Text -match "^[0-9]{2}[X0-9]{1,4}[A-D]?$"){
                $inst_sequens_numbers += $word
                
            }
         }        
        
        
         
        foreach($word in $page_words){



            # if ($word.Text -match 'PRVXV') {
            #     $word 
            # }
            # if ($word.Text -match '^540001$') {
            #     $word 
            # }
            
            
            
            # continue


            $offset = 10
		    if($word.Text -match "^(" + $Inst_Regexes+ ")$" ){
                # $word.Text
                if($word.TextOrientation -eq "Horizontal"){
                    foreach($seq_word in $inst_sequens_numbers){
                        if ($word.BoundingBox.Centroid.Y - $offset  -ge $seq_word.BoundingBox.Bottom -and 
                            $word.BoundingBox.Centroid.Y - $offset  -le $seq_word.BoundingBox.Top -and 
                            $word.BoundingBox.Centroid.X -ge $seq_word.BoundingBox.Left -and 
                            $word.BoundingBox.Centroid.X -le $seq_word.BoundingBox.Right){
                                $word.Text + "-" + $seq_word.Text
                            $record = New-Object Tag2Doc
                            $record.Tag_number = $Tag_number_Prefix + "-" + $word.Text + "-" + $seq_word.Text
                            $record.short_id = $seq_word.Text
                            $record.document_number = $file.BaseName 
                            $record.ST = "Advanced Instrument Tag Search"
                            $record.date = $date
                            $record.doc_date = $revision_date
                            $record.issue_reason = $reaseon_for_issue
                            $tags += $record
                            # $File_data += $record
                            }
                        }
                    }
                if($word.TextOrientation -eq "Horizontal"){
                    foreach($seq_word in $inst_sequens_numbers){
                        if ($word.BoundingBox.Centroid.Y + $offset  -ge $seq_word.BoundingBox.Bottom -and 
                            $word.BoundingBox.Centroid.Y - $offset  -le $seq_word.BoundingBox.Top -and 
                            $word.BoundingBox.Centroid.X -ge $seq_word.BoundingBox.Left -and 
                            $word.BoundingBox.Centroid.X -le $seq_word.BoundingBox.Right){
                                $word.Text + "-" + $seq_word.Text
                            $record = New-Object Tag2Doc
                            $record.Tag_number = $Tag_number_Prefix + "-" + $word.Text + "-" + $seq_word.Text
                            $record.short_id = $seq_word.Text
                            $record.document_number = $file.BaseName 
                            $record.ST = "Advanced Instrument Tag Search"
                            $record.date = $date
                            $record.doc_date = $revision_date
                            $record.issue_reason = $reaseon_for_issue
                            $tags += $record
                            # $File_data += $record
                            }
                        }
                    }
                
                elseif($word.TextOrientation -eq "Rotate270"){
                    foreach($seq_word in $inst_sequens_numbers){
                    if($word.BoundingBox.Centroid.Y -ge $seq_word.BoundingBox.Bottom -and 
                        $word.BoundingBox.Centroid.Y -le $seq_word.BoundingBox.Top -and 
                        $word.BoundingBox.Centroid.X + $offset  -ge $seq_word.BoundingBox.Left -and 
                        $word.BoundingBox.Centroid.X + $offset  -le $seq_word.BoundingBox.Right){
                        $record = New-Object Tag2Doc
                        $record.Tag_number = $Tag_number_Prefix + "-" + $word.Text + "-" + $seq_word.Text
                        $record.short_id = $seq_word.Text
                        $record.document_number = $file.BaseName 
                        $record.ST = "Advanced Instrument Tag Search"
                        $record.date = $date
                        $record.doc_date = $revision_date
                        $record.issue_reason = $reaseon_for_issue
                        $tags += $record
                        # $File_data += $record
                        }
                    }
                }
                elseif($word.TextOrientation -eq "Rotate90"){
                    foreach($seq_word in $inst_sequens_numbers){
                    if($word.BoundingBox.Centroid.Y -ge $seq_word.BoundingBox.Bottom -and 
                        $word.BoundingBox.Centroid.Y -le $seq_word.BoundingBox.Top -and 
                        $word.BoundingBox.Centroid.X + $offset  -ge $seq_word.BoundingBox.Left -and 
                        $word.BoundingBox.Centroid.X + $offset  -le $seq_word.BoundingBox.Right){
                        $record = New-Object Tag2Doc
                        $record.Tag_number = $Tag_number_Prefix + "-" + $word.Text + "-" + $seq_word.Text
                        $record.short_id = $seq_word.Text
                        $record.document_number = $file.BaseName 
                        $record.ST = "Advanced Instrument Tag Search"
                        $record.date = $date
                        $record.doc_date = $revision_date
                        $record.issue_reason = $reaseon_for_issue
                        $tags += $record
                        # $File_data += $record
                        }
                    }
                }
                else{
                    #ECHO "Uknown rotation detected"
                    $mess = $word.Text + '- Rotation: ' + $word.TextOrientation
                    Write-Host "[Warning] " -ForegroundColor Yellow -NoNewline
                    Write-Host "Cannot capture orientation for: $mess"

                }
            }        
        }
    
    
    # for ($i = 0; $i -lt $tags.Count; $i++) {
    #     $tags[$i] = $tags[$i] +';'+ $date
    # }

    # $tags | Export-Csv -Path $tag_report -Append -NoTypeInformation -NoClobber
    # $tags |Select-Object -Unique -Property Tag_number | Export-Csv -Path $tag_report -Append -NoTypeInformation -NoClobber
    
    #ECHO $watch.Elapsed.TotalSeconds
    #$watch.reset()
    }
    # $temp_file = [System.IO.Path]::Combine($temp_folder , $file.BaseName +".csv")
    # $File_data = $File_data | Where-Object {-not([string]::IsNullOrEmpty($_.Tag_number)) }
    # $File_data | Export-Csv -Path $temp_file -NoTypeInformation -Encoding UTF8
}
# $tags = $tags | Where-Object {-not([string]::IsNullOrEmpty($_.Tag_number)) }

# $tags | Export-Csv -Path $tag_report -NoTypeInformation -Encoding UTF8
# $tags | Select-Object -Property @{Name = 'ReferencedId'; Expression = {$_.Tag_number}}, 
# @{Name = "Document_ID"; Expression = {$_.Document_number}}, 
# @{Name = 'Alias'; Expression = {$_.Tag_number}} | Export-Csv -Path $Tag_to_Doc -NoTypeInformation -Delimiter ','
# $tags | Export-Csv -Path 'W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Temp\test.csv' -NoTypeInformation -Delimiter ','
$tags | Select-Object -Property Tag_number