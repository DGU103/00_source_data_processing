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

$files_Dir = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\CPPR1-MDM5-ASBJA-10-Q68031-0001\03"

$date = Get-Date -Format 'dd/MM/yyyy'

$Inst_Regexes = @("AAH","AAHH","AI","AIS","AIT","APB","AR","ARC","ASP","AT","BDV","BX","BY","CAM","CMO","CMP","CPF","CPJ","CPR","CS","CVA","CY","DI","DT","EPB","ESDV","EWS","EX","EY","FAH","FAHH","FAL","FALL","FC","FCV","FE","FG","FI","FIT","FIV","FMX","FO","FPS","FQ","FQI","FQV","FQVY","FS","FSH","FSHH","FSL","FSLL","FT","FVI","FX","FY","GD","GDAH","GDAHH","GDR","GDS","GDT","GVA","GVAA","HC","HD","HDAH","HDAHH","HDC","HDR","HDS","HDT","HF","HG","HGAH","HGAHH","HGS","HIT","HR","HRAH","HS","HSS","HT","HVA","IAM","ICD","ID","IMS","IPC","IR","IRAH","JBC","JBE","JBF","JBJ","JBS","LAH","LAHH","LAL","LALL","LC","LCV","LG","LI","LIT","LOS","LRS","LS","LSC","LSD","LSH","LSHH","LSHL","LSL","LSLL","LSS","LT","LVI","LY","MAC","MACA","MCT","MI","MOV","MRD","MT","MWS","OCP","OWS","PA","PAH","PAHH","PAL","PALL","PB","PC","PCD","PCV","PDAH","PDAHH","PDAL","PDALL","PDC","PDCV","PDI","PDIT","PDRC","PDS","PDSH","PDSHH","PDSL","PDSLL","PDT","PDY","PE","PI","PIT","PRI","PRV","PS","PSE","PSH","PSHH","PSL","PSLL","PSV","PT","PV","PVI","PX","PY","R","RCU","RD","RTD","RTU","SAH","SAHH","SAL","SALL","SD","SDAH","SDV","SE","SI","SL","SOV","SS","SSH","SSL","ST","SVC","SVP","SWS","SX","SY","TAH","TAHH","TAL","TALL","TC","TCV","TDAH","TDAL","TDIC","TDY","TE","TES","TI","TIT","TMX","TS","TSH","TSHH","TSHL","TSL","TSLL","TSV","TT","TVI","TW","TY","UA","UV","VAH","VAHH","VDU","VGDAH","VGDAHH","VHDAH","VHDAHH","VHGAH","VHGAHH","VHRAH","VIRAH","VMACA","VSDAH","VT","WAA","WMA","WMH","WML","WMR","WMV","WT","X","XA","XAH","XAHH","XC","XCT","XCV","XEP","XI","XL","XPI","XPS","XS","XT","XY","Y","YSL","ZAH","ZAHH","ZE","ZI","ZIC","ZIO","ZL","ZLC","ZLO","ZS","ZSC","ZSO","ZT", "HCV", "XZSL", "PRVXV") -join "|"


$files = Get-ChildItem -Path $files_Dir -Filter *.pdf -Recurse #| Where-Object {$_.BaseName -match $document_selection_criteria}

$tags = @([Tag2Doc]::new())

$rotattion_list = @()
$all_text = @()
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
        $all_text += $page_words
        $scale = 1
        $h = $page.CropBox.Bounds.Height
        $w = $page.CropBox.Bounds.Width
        if ($h > $w) {
            $ph = $w
        }
        else{
            $ph = $h
        }
        # ( 594 x 841)
        $scale = 594 / $ph

        # $page.Rotation
        # Write-Host "H: $h x W: $w"
        # if($page.Size -notmatch "A1"){
        #     CONTINUE
        # }
        if ($file.FullName -notmatch '-10-') {
            CONTINUE
        }
        foreach($word in $page_words){
            if ($word.Text -match "^[0-9]{6}[A-Z]?$" -or $word.Text -match "^[0-9]{2}[X0-9]{1,4}[A-D]?$"){
                $inst_sequens_numbers += $word
            }
         }        
         
        foreach($word in $page_words){
            if($word -match 'ZI|502101' -and $word.TextOrientation -eq 'Rotate270'){
                # $word.Text
                # $word.BoundingBox.Centroid
                # $word.BoundingBox.Bottom
                # $word.BoundingBox.Top

            }
            $offset = 0
		    if($word.Text -match "^(" + $Inst_Regexes + ")$" ){
                if($word.TextOrientation -eq "Horizontal"){
                    foreach($seq_word in $inst_sequens_numbers){
                        if ($word.BoundingBox.Centroid.Y + $offset  -ge $seq_word.BoundingBox.Bottom -and 
                            $word.BoundingBox.Centroid.Y + $offset  -le $seq_word.BoundingBox.Top -and 
                            $word.BoundingBox.Centroid.X -ge $seq_word.BoundingBox.Left -and 
                            $word.BoundingBox.Centroid.X -le $seq_word.BoundingBox.Right){
                                # $word.Text + "-" + $seq_word.Text
                                # $record = New-Object Tag2Doc
                                # $record.Tag_number = $Tag_number_Prefix + "-" + $word.Text + "-" + $seq_word.Text
                                # $record.short_id = $seq_word.Text
                                # $record.document_number = $file.BaseName 
                                # $record.ST = "Advanced Instrument Tag Search"
                                # $record.date = $date
                                # $record.doc_date = $revision_date
                                # $record.issue_reason = $reaseon_for_issue
                                # $tags += $record
                            }
                        }
                        
                    }
                elseif($word.TextOrientation -eq "Rotate270"){
                    foreach($seq_word in $inst_sequens_numbers){
                        if($word.BoundingBox.Centroid.Y -ge ($seq_word.BoundingBox.BottomLeft.Y ) -and 
                            $word.BoundingBox.Centroid.Y -le ($seq_word.BoundingBox.TopRight.Y) -and 
                            $word.BoundingBox.Centroid.X  + $offset -ge ($seq_word.BoundingBox.BottomLeft.X ) -and 
                            $word.BoundingBox.Centroid.X  + $offset -le ($seq_word.BoundingBox.TopRight.X )){
                                $record = New-Object Tag2Doc
                                $record.Tag_number = $Tag_number_Prefix + "-" + $word.Text + "-" + $seq_word.Text
                                # $record.short_id = $seq_word.Text
                                # $record.document_number = $file.BaseName 
                                # $record.ST = "Advanced Instrument Tag Search"
                                # $record.date = $date
                                # $record.doc_date = $revision_date
                                # $record.issue_reason = $reaseon_for_issue
                                # $tags += $record
                                $record.Tag_number
                                Write-Host '_RECT'
                                $word.BoundingBox.BottomLeft.X.ToString() + ',' + $word.BoundingBox.BottomLeft.Y.ToString()
                                '@' + $word.BoundingBox.TopRight.X.ToString() + ',' + $word.BoundingBox.TopRight.Y.ToString()
                                Write-Host '_RECT'
                                $seq_word.BoundingBox.BottomLeft.X.ToString() + ',' + $seq_word.BoundingBox.BottomLeft.Y.ToString()
                                '@' + $seq_word.BoundingBox.TopRight.X.ToString() + ',' + $seq_word.BoundingBox.TopRight.Y.ToString()

                                

                            }

                        }
                        

                }
                elseif($word.TextOrientation -eq "Rotate90"){
                    foreach($seq_word in $inst_sequens_numbers){
                        if($word.BoundingBox.Centroid.Y -ge $seq_word.BoundingBox.Bottom -and 
                            $word.BoundingBox.Centroid.Y -le $seq_word.BoundingBox.Top -and 
                            $word.BoundingBox.Centroid.X + $offset  -ge $seq_word.BoundingBox.Left -and 
                            $word.BoundingBox.Centroid.X + $offset  -le $seq_word.BoundingBox.Right){
                                # $record = New-Object Tag2Doc
                                # $record.Tag_number = $Tag_number_Prefix + "-" + $word.Text + "-" + $seq_word.Text
                                # $record.short_id = $seq_word.Text
                                # $record.document_number = $file.BaseName 
                                # $record.ST = "Advanced Instrument Tag Search"
                                # $record.date = $date
                                # $record.doc_date = $revision_date
                                # $record.issue_reason = $reaseon_for_issue
                                # $tags += $record
                            }
                    }
                    

                }
                elseif($word.TextOrientation -eq "Rotate180"){
                    

                }
                elseif($word.TextOrientation -eq "Other"){
                    

                }
                else{
                    

                }

            }        
        }
    }
}
# $tags | Select-Object -Property Tag_number
$rotattion_list = $rotattion_list | Select-Object -Unique
# Write-Host $rotattion_list -ForegroundColor Green -BackgroundColor Black
# $tags | Export-Csv -Path 00_source_data_processing\00_Document_Indexing\NOT_IN_USE_99_Doc_indexing_Research.txt
# $all_text | Out-File -FilePath 00_source_data_processing\00_Document_Indexing\NOT_IN_USE_99_Doc_indexing_Research.txt