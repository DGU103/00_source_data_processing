<#

    Purpose: Parallel indexing of PDF files for a given EPC package.
             The script scans PDF files for bookmarks and text (using UglyToad.PdfPig),
             and produces Tag2Doc and Doc2Doc CSV report.
#>

param (
    [Parameter(Mandatory=$true)]
    #[ValidateSet('05','06','11','12','13')]
    [string]$epc,
    [switch]$EnableDebug
)


if ($EnableDebug.IsPresent) {
    $global:DEBUG_ENABLED = $true
} else {
    $global:DEBUG_ENABLED = $false
}

#Manual Trigger for ON/OFF Debugger
# $DEBUG_ENABLED = $true

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "INDEXING"
$finished = $false

#Load common file
. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "Running $scriptname for EPCIC $epc"
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"

if($global:DEBUG_ENABLED) {
    Write-Log -Level INFO -Message "DEBUG logging is ENABLED."
} else {
    Write-Log -Level INFO -Message "DEBUG logging is DISABLED."
}


# Determine root folder based on EPC value
if($epc -in @('11','12','13')) {
   $root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\"
   #$root_path = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\Indexing\smallbatch\"
}
elseif($epc -eq '06') {
    $root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\CPP03\Source\Indexing\"
}
elseif($epc -eq '05') {
    $root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\"
}

$local_path = $PSScriptRoot
$files_Dir = Join-Path $root_path ("EPC" + $epc + "_Source")
#$files_Dir = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\Indexing\smallbatch"
$tag_report = Join-Path $root_path ("EPCIC" + $epc + "_indexing_report.csv")
#$tag_report = "H:\My Documents\Artifacts\Indexing\out\EPCIC12_Tagsiki.csv"
#$doc_report = "H:\My Documents\Artifacts\Indexing\out\EPCIC12_Docsiki.csv"
$doc_report = Join-Path $root_path ("EPCIC" + $epc + "DOC_indexing_report.csv")

Write-Log -Level INFO -Message "Collecting PDF files from: $files_Dir"

$inArray = Get-ChildItem -Path $files_Dir -Filter '*.pdf' -Recurse
if(-not $inArray -or $inArray.Count -eq 0) {
   Write-Log -Level WARN -Message "No PDF files found in $files_Dir for EPC $epc. Exiting."
    return
}
Write-Log -Level INFO -Message "Found $($inArray.Count) PDFs to process."

################################################################################
# --- Split Files into Batches for Parallel Processing ---
################################################################################
$parts = 8
[int] $partSize = [Math]::Round($inArray.Count / $parts, 0)
if($partSize -eq 0) {
    Write-Log -Level ERROR -Message "$parts sub-arrays requested, but only $($inArray.Count) files found."
    throw "$parts sub-arrays requested, but only $($inArray.Count) files found."
}
$extraSize = $inArray.Count - ($partSize * $parts)
$offset = 0
$jobs_list = @()

foreach ($i in 1..$parts) {
    $temp = [System.IO.FileInfo[]]$inArray[$offset..($offset + $partSize + ([bool]$extraSize) - 1)]

    $job_id = "EPC${epc}_Batch${i}_Indexing"
    Write-Log -Level INFO -Message "Starting job '$job_id' with $($temp.Count) PDFs."

    $job = Start-Job -Name $job_id -ScriptBlock {
        param($files, $local_path, $epc, $global:DEBUG_ENABLED, $global:scriptname)


        class Tag2Doc {
            [string] $Tag_number
            [string] $Document_number
            [string] $doctitle
            [string] $ST
            [string] $DATE
            [string] $doc_date
            [string] $issue_reason
            [string] $file_full_path
            [string] $doctype
            [string] $issuance_code
        }
        class Doc2Doc {
            [string] $source_doc_id
            [string] $ref_doc_id
            [string] $DATE
        }

        $filetimes = @()

        . "$using:PSScriptRoot\..\Common_Functions.ps1"

      
        $dll_path = [System.IO.Path]::Combine($local_path, 'lib', 'UglyToad.PdfPig.dll')
        Import-Module $dll_path -ErrorAction SilentlyContinue
        $dll_path = [System.IO.Path]::Combine($local_path, 'lib', 'UglyToad.PdfPig.DocumentLayoutAnalysis.dll')
        Import-Module $dll_path -ErrorAction SilentlyContinue
        $dll_path = [System.IO.Path]::Combine($local_path, 'lib', 'BouncyCastle.Crypto.dll')
        Import-Module $dll_path -ErrorAction SilentlyContinue
        $dll_path = [System.IO.Path]::Combine($local_path, 'lib', 'itextsharp.dll')
        Import-Module $dll_path -ErrorAction SilentlyContinue
        $dll_path = [System.IO.Path]::Combine($local_path, 'lib', 'itextsharp.pdfa.dll')
        Import-Module $dll_path -ErrorAction SilentlyContinue

        $date = Get-Date -Format 'dd/MM/yyyy'

        $Inst_Regexes = @("AAH","AAHH","AI","AIS","AIT","APB","AR","ARC","ASP","AT","AV","BDV","BLD","BPR","BX","BY","CAM","CC","CHV","CI","CMO","CMP","CPF","CPJ","CPR","CS","CTP","CVA","CY","DI","DRS","DT","EJX","EPB","EPR","ESDV","EWS","EX","EY","FA","FAH","FAHH","FAL","FALL","FC","FCS","FCV","FE","FG","FHA","FI","FIT","FIV","FMX","FO","FPS","FQ","FQI","FQV","FQVY","FS","FSH","FSHH","FSL","FSLL","FT","FVI","FVS","FX","FY","GD","GDAH","GDAHH","GDR","GDS","GDT","GLV","GVA","GVAA","HC","HCS","HCV","HD","HDAH","HDAHH","HDC","HDR","HDS","HDT","HF","HG","HGAH","HGAHH","HGS","HIT","HR","HRAH","HS","HSS","HT","HVA","HVS","IAM","ICD","ID","ILK","IMS","IPC","IR","IRAH","JBC","JBE","JBF","JBJ","JBS","LAH","LAHH","LAL","LALL","LC","LCV","LG","LI","LIT","LOS","LRS","LS","LSC","LSD","LSH","LSHH","LSHL","LSL","LSLL","LSS","LT","LVI","LY","MAC","MACA","MCT","MCV","MI","MOV","MRD","MT","MWS","OCP","OWS","PA","PAH","PAHH","PAL","PALL","PB","PC","PCD","PCV","PDAH","PDAHH","PDAL","PDALL","PDC","PDCV","PDI","PDIT","PDRC","PDS","PDSH","PDSHH","PDSL","PDSLL","PDT","PDY","PE","PI","PIT","PRI","PRS","PRV","PS","PSE","PSH","PSHH","PSL","PSLL","PSV","PT","PV","PVI","PX","PY","R","RCU","RD","RO","RTD","RTU","S","SAH","SAHH","SAL","SALL","SCP","SD","SDAH","SDV","SE","SI","SL","SOV","SPR","SS","SSH","SSL","SSSV","ST","SVC","SVP","SWS","SX","SY","TAH","TAHH","TAL","TALL","TC","TCV","TDAH","TDAL","TDIC","TDY","TE","TES","TI","TIT","TMX","TS","TSH","TSHH","TSHL","TSL","TSLL","TSV","TT","TVI","TW","TY","UA","UV","VAH","VAHH","VDU","VGDAH","VGDAHH","VHDAH","VHDAHH","VHGAH","VHGAHH","VHRAH","VIRAH","VMACA","VPSV","VSDAH","VT","WAA","WCV","WMA","WMH","WML","WMR","WMV","WT","X","XA","XAH","XAHH","XC","XCT","XCV","XEP","XI","XL","XPI","XPS","XS","XT","XY","Y","YSL","ZAH","ZAHH","ZE","ZI","ZIC","ZIO","ZL","ZLC","ZLO","ZS","ZSC","ZSO","ZT","2WV","3WV","AAL","ABE","ABI","ABIT","ABT","AC","ACUSH","ACUSL","ADTN","AEN","AGE","AGI","AGIT","AGT","AH","AHA","AHE","AHI","AHIT","AHS","AHT","AIC","AIN","AITN","AME","AMI","AMIT","AMT","AO","AOJ","AP","APCSH","APCSL","ART","ARV","ASCSH","ASCSL","ASH","ASHL","ASL","ASPSH","ASPSL","ATM","ATN","AVY","AWT","AX","AY","AZ","AZE","AZI","AZIT","AZL","AZP","AZR","AZS","AZSC","AZSO","AZT","AZTN","BAL","BDIM","BDIOM","BDOM","BFCL","BFV","BG","BI","BIAD","BIAL","BIALS","BIAS","BIC","BIT","BL","BP","BPV","BR","BRC","BRT","BS","BSG","BSH","BSHL","BSL","BSP","BT","BTC","BTF","BTH","BTHA","BTHL","BTHLI","BTHLR","BTK","BTKL","BTM","BW","BZE","BZL","BZP","BZR","BZS","BZW","CAH","CAL","CE","CGE","CGT","CGTN","CIT","CSC","CSH","CSL","CSO","DE","DHS","DO","DPSH","DPSL","DR","DTT","DX","DY","DZY","EE","EG","EI","EIC","EIT","ER","ERC","ERT","ESD","ESH","ESHL","ESL","ET","EZE","EZI","EZL","EZP","EZR","EZW","FBE","FBI","FBIT","FBS","FBT","FDTN","FEN","FFC","FFE","FFI","FFIC","FFR","FFRC","FFSH","FFSL","FGE","FGI","FGIT","FGN","FGS","FGT","FHE","FHG","FHI","FHIT","FHS","FHT","FICV","FIN","FITN","FITNSH","FITNSL","FJB","FL","FM","FME","FMI","FMIT","FMS","FMT","FQE","FQG","FQIC","FQIT","FQR","FQRC","FQSH","FQSL","FQT","FQX","FQY","FR","FRC","FRT","FSHL","FSV","FTN","FU","FV","FVY","FZ","FZE","FZI","FZIT","FZL","FZP","FZR","FZS","FZSC","FZSO","FZT","FZTN","FZU","FZV","FZW","FZY","FZZS","FZZSC","FZZSO","GAE","GAT","GATN","GAV","GBE","GBT","GBTN","GCE","GCFT","GCFTN","GCT","GCTN","GDE","GDTN","GEE","GET","GETN","GFE","GFT","GFTN","GGAE","GGAT","GGATN","GGBE","GGBT","GGBTN","GGCE","GGCT","GGCTN","GGDE","GGDT","GGDTN","GGE","GGEE","GGET","GGETN","GGFE","GGFT","GGFTN","GGGE","GGGT","GGGTN","GGHE","GGHT","GGHTN","GGIE","GGIT","GGITN","GGJE","GGJT","GGJTN","GGKE","GGKT","GGKTN","GGLE","GGLT","GGLTN","GGME","GGMT","GGMTN","GGNE","GGNT","GGNTN","GGOT","GGOTN","GGPT","GGPTN","GGRT","GGRTN","GGST","GGSTN","GGT","GGTN","GHCT","GHCTN","GHE","GHLT","GHLTN","GHT","GHTN","GIE","GIT","GITN","GJE","GJT","GJTN","GKE","GKT","GKTN","GLE","GLT","GLTN","GMCE","GMCT","GMCTN","GME","GMT","GMTN","GNE","GNOE","GNOT","GNOTN","GNT","GNTN","GOE","GOT","GOTN","GPE","GPOT","GPOTN","GPT","GPTN","GRE","GRT","GRTN","GSE","GST","GSTN","GTT","GTTN","GUE","GUT","GUTN","GVE","GVT","GVTN","GWE","GWT","GWTN","GXE","GXT","GXTN","GYE","GYT","GYTN","GZE","GZT","GZTN","HBB","HBS","HBV","HBY","HHS","HIC","HL","HOA","HSA","HSC","HSO","HSZ","HV","HVY","HX","HY","HYC","HYO","HZIC","HZIO","HZR","HZS","HZSC","HZSO","HZV","HZY","HZZS","HZZSC","HZZSO","IA","II","IIC","IIT","IL","IP","IRC","IRT","ISH","ISHL","ISL","IT","IV","IX","IY","IZE","IZI","IZL","JE","JG","JI","JIC","JIT","JO","JOI","JOR","JQX","JR","JRC","JRT","JSH","JSHL","JSL","JT","JX","JY","JZE","JZI","JZL","JZP","JZR","K","KC","KCV","KE","KG","KI","KIC","KIT","KL","KME","KOG","KOL","KOX","KPE","KQI","KR","KRC","KRT","KSH","KSHL","KSL","KT","KV","KX","KY","KZSC","KZSD","KZSO","KZSS","KZT","KZV","KZY","KZZSC","KZZSO","LA","LACT","LAD","LAHI","LBE","LBI","LBIT","LBP","LBS","LBT","LCI","LCVI","LD","LDI","LDIN","LDIT","LDT","LDTN","LDZT","LDZU","LEI","LEN","LGE","LGI","LGIT","LGS","LGT","LHE","LHG","LHI","LHIT","LHS","LHT","LICI","LIN","LINSH","LINSL","LITN","LITNSH","LITNSL","LL","LLC","LLG","LME","LMI","LMIT","LMS","LMT","LO","LR","LRC","LSHI","LSLI","LTI","LTN","LUX","LV","LVS","LVY","LW","LWG","LX","LZE","LZI","LZIN","LZIT","LZL","LZP","LZS","LZSC","LZSH","LZSL","LZSN","LZSO","LZT","LZTN","LZTNSH","LZTNSL","LZV","LZY","MA","MBOV","MBZSC","MBZSO","MCU","MHHS","MHS","MHT","MLPSH","MLPSL","MZHS","MZLC","MZLO","MZSC","MZSO","MZZS","MZZSC","MZZSO","MZZT","NO","NOC","OM","OV","OWD","OX","PAD","PBI","PBIT","PBS","PBT","PCVN","PDBI","PDBIT","PDBS","PDBT","PDE","PDG","PDGI","PDGIT","PDGN","PDGS","PDGT","PDHCV","PDHG","PDHI","PDHIT","PDHS","PDHSH","PDHSL","PDHT","PDIN","PDITN","PDMI","PDMIT","PDMS","PDMT","PDR","PDRT","PDTN","PDV","PDVY","PDX","PDZI","PDZIT","PDZS","PDZSC","PDZSH","PDZSO","PDZT","PDZTN","PEN","PFL","PFR","PFX","PG","PGI","PGIT","PGN","PGS","PHG","PHI","PHIT","PHS","PHT","PHV","PIN","PITN","PJR","PK","PKL","PKR","PKX","PLC","PLPSH","PLPSL","PMI","PMIT","PMS","PMT","PN","POL","PP","PRC","PRT","PSHL","PSVNSH","PSVNSL","PTC","PTN","PVN","PVY","PZE","PZI","PZIN","PZIT","PZL","PZP","PZS","PZSC","PZSH","PZSL","PZSO","PZT","PZTN","PZTNSH","PZTNSL","PZV","PZY","PZZS","PZZSC","PZZSO","QE","QI","QIC","QIT","QOL","QQ","QQI","QQR","QQX","QR","QRC","QRT","QSH","QSHL","QSL","QT","QY","RAV","RE","REG","RI","RIC","RIT","ROR","ROX","RP","RQ","RQI","RQL","RRC","RRT","RSH","RSHL","RSL","RT","RW","RY","RZE","RZL","RZP","SB","SC","SCN","SCNSH","SCNSL","SCV","SG","SHT","SIC","SIT","SJ","SLO","SME","SMT","SP","SR","SRC","SRT","SRV","SSHL","SSV","SV","SVY","SZE","SZIT","SZT","SZY","TBE","TBI","TBIT","TBS","TBT","TD","TDA","TDC","TDCV","TDE","TDG","TDI","TDIT","TDL","TDR","TDRC","TDRT","TDS","TDSH","TDSL","TDT","TDTN","TDV","TDVY","TDX","TEN","TFI","TFR","TFX","TG","TGE","TGI","TGIT","TGS","TGT","THCV","THE","THG","THI","THIT","THS","THT","THV","TIN","TINSH","TINSL","TIS","TISHL","TITN","TJ","TJE","TJX","TK","TKR","TKX","TL","TME","TMI","TMIT","TMS","TMT","TOR","TP","TR","TRC","TRS","TRT","TSE","TTN","TV","TVY","TZE","TZI","TZIT","TZL","TZP","TZR","TZS","TZSC","TZSH","TZSL","TZSO","TZT","TZTN","TZV","TZW","TZY","UE","UI","UJ","UJR","UL","ULO","UR","USD","UTN","UVY","UX","UY","UZ","UZE","UZL","UZP","UZR","UZS","UZSC","UZSO","UZW","UZY","VBE","VBIT","VBS","VBT","VE","VFD","VG","VGE","VGIT","VGS","VGT","VHE","VHIT","VHS","VHT","VI","VIT","VL","VME","VMIT","VMS","VMT","VP","VR","VRT","VS","VTA","VX","VXE","VXG","VXI","VXL","VXME","VXR","VXT","VXX","VY","VYE","VYME","VYP","VYT","VZ","VZE","VZG","VZIT","VZL","VZP","VZR","VZS","VZT","VZX","WA","WAI","WAR","WAX","WC","WDC","WDCV","WDI","WDIC","WDIT","WDL","WDR","WDRC","WDRT","WDSH","WDSL","WDT","WDX","WE","WEC","WEN","WFI","WFR","WFX","WG","WHS","WI","WIC","WIN","WIT","WKI","WKR","WKX","WL","WMS","WQI","WQL","WQR","WQX","WR","WRT","WS","WSH","WSHL","WSL","WTN","WUL","WUP","WUR","WX","WXE","WXI","WXL","WXR","WXX","WY","WYE","WYL","WZ","WZE","WZIT","WZL","WZR","WZS","WZT","WZX","WZXL","WZXP","WZXR","WZYL","WZYP","WZYR","XBY","XG","XHY","XHZSC","XIO","XLC","XLO","XP","XQ","XR","XV","XVN","XVY","XW","XX","XYC","XYD","XYH","XYL","XYO","XYS","XZC","XZIC","XZIO","XZLC","XZLO","XZO","XZS","XZSC","XZSD","XZSH","XZSL","XZSO","XZSP","XZSS","XZV","XZY","XZYC","XZYO","XZZLC","XZZLO","XZZS","XZZSC","XZZSD","XZZSH","XZZSL","XZZSO","XZZSS","YC","YE","YI","YIC","YL","YR","YSH","YT","YX","YY","YZ","YZE","YZL","YZR","YZX","ZC","ZCV","ZDC","ZDCV","ZDE","ZDG","ZDI","ZDIC","ZDIT","ZDL","ZDR","ZDRC","ZDRT","ZDSH","ZDSL","ZDT","ZDX","ZDXE","ZDXG","ZDXI","ZDXL","ZDXR","ZDXX","ZDY","ZDYE","ZDYG","ZDYI","ZDYL","ZDYR","ZDYX","ZDZE","ZDZI","ZDZL","ZDZR","ZDZX","ZG","ZHS","ZHSC","ZHSO","ZHT","ZIT","ZME","ZMT","ZO","ZP","ZR","ZRC","ZRT","ZSD","ZSE","ZSH","ZSHL","ZSL","ZSP","ZST","ZTN","ZUI","ZUL","ZUP","ZUR","ZW","ZX","ZXE","ZXG","ZXI","ZXL","ZXR","ZY","ZYE","ZYG","ZYL","ZYR","ZYX","ZZ","ZZI","ZZX","ZZXE","ZZXI","ZZXL","ZZXP","ZZXR","ZZYE","ZZYI","ZZYL","ZZYP","ZZYR") -join "|"
       $Light_Regex = Import-Csv -Delimiter ";" -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv"

        $tags = @()
        $docs = @()

        foreach ($f in $files) {


            $file = [System.IO.FileInfo]::new($f.FullName)

            $sw = [System.Diagnostics.Stopwatch]::StartNew()

            # If _null.xml is missing, log error and skip file.
            $metaPath = $file.FullName.Replace('.pdf', '_null.xml')
            if (-not (Test-Path $metaPath)) {
                Write-Log -Level ERROR -Message "Metadata for $($file.FullName) not found. Skipping file."
                continue
            }

            # Read XML safely.
            try {
                [xml]$XmlDocument = Get-Content $metaPath
            }
            catch {
                Write-Log -Level ERROR -Message "Unable to read XML for $($file.FullName): $($_.Exception.Message). Skipping."
                continue
            }

            $XmlNamespace = @{ a = "http://www.aveva.com/VNET/eiwm" }
            $revision_date_item = Select-Xml -Xml $XmlDocument -XPath "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision_date']/a:Value" -Namespace $XmlNamespace
            $revision_date = ($revision_date_item.Node.'#text' -split " ")[0]

            $doctype = (Select-Xml -Xml $XmlDocument -XPath "//a:Template/a:Object/a:Characteristic[a:Name='pjc_doc_type']/a:Value" -Namespace $XmlNamespace).Node.InnerText
            
            $doctitle = (Select-Xml -Xml $XmlDocument -XPath "//a:Template/a:Object/a:Characteristic[a:Name='title']/a:Value" -Namespace $XmlNamespace).Node.InnerText

            $issuance_code = (Select-Xml -Xml $XmlDocument -XPath "//a:Template/a:Object/a:Characteristic[a:Name='pjc_last_return_code']/a:Value" -Namespace $XmlNamespace).Node.InnerText

            $reason_item = Select-Xml -Xml $XmlDocument -XPath "//a:Template/a:Object/a:Characteristic[a:Name='pjc_revision_object']/a:Value" -Namespace $XmlNamespace
            $reasonText = $reason_item.Node.'#text'
            if ($reasonText -match "CLD") {
                Write-Log -Level WARN -Message "Skipping $($file.Name) because reason is CLD."
                continue
            }
            

            # Process bookmarks with iTextSharp
            try {
                $PdfReader = New-Object iTextSharp.text.pdf.PdfReader($file.FullName)
                $BookMarks = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($PdfReader)
  
            #'Get-Bookmarks' Function inside of the 'Common Functions.ps1'

            Get-Bookmarks -Bookmarks $BookMarks -Light_Regex $Light_Regex -fileBaseName $file.BaseName -date $date -revision_date $revision_date -reasonText $reasonText -fileFullName $file.FullName -doctype $doctype -doctitle $doctitle -issuance_code $issuance_code -tags ([ref]$tags)

            }

            catch {
                Write-Log -Level ERROR -Message "Cannot process bookmarks in $($file.FullName): $($_.Exception.Message)"
            }

            # Build a prefix from the file name
            $spl = $file.Name -split "-"
            if($spl.Count -lt 3) {
                $Tag_number_Prefix = "XXXXA"
            }
            else {
                $Tag_number_Prefix = $spl[2].Substring(0,4) + "A"
            }

            # Open PDF with UglyToad.PdfPig
            try {
                $pdf = [UglyToad.PdfPig.PdfDocument]::Open($file.FullName)
            }
            catch {
                Write-Log -Level WARN -Message "UglyToad PdfPig failed to open $($file.FullName): $($_.Exception.Message). Skipping text extraction."
                continue
            }
            if(-not $pdf) { continue }

            for($p = 1; $p -le $pdf.NumberOfPages; $p++){
                $page = $null
                try {
                    $page = $pdf.GetPage($p)
                }
                catch {
                    Write-Log -Level WARN -Message "Cannot read page #$p from $($file.FullName): $($_.Exception.Message). Skipping page."
                    continue
                }
                if(-not $page) { continue }

                $page_words = $null
                try {
                    $page_words = $page.GetWords([UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor]::Instance)
                }
                catch {
                    Write-Log -Level WARN -Message "Error extracting words from page #$p of $($file.Name): $($_.Exception.Message)"
                    continue
                }
                if(-not $page_words) { continue }

                # For each word on the page, check against Light_Regex patterns
                foreach($word in $page_words){
                    if([string]::IsNullOrEmpty($word.Text)) { continue }
                    foreach($regex in $Light_Regex) {
                        if($word.Text -match $regex.Regexp){
                            $record = [Tag2Doc]::new()
                            $record.Tag_number = $word.Text
                            $record.Document_number = $file.BaseName
                            $record.ST = $regex.Naming_template_ID
                            $record.DATE = $date
                            $record.doc_date = $revision_date
                            $record.issue_reason = $reasonText
                            $record.file_full_path = $file.FullName
                            $record.doctype = $doctype
                            $record.doctitle = $doctitle
                            $record.issuance_code = $issuance_code
                            $tags += $record
                            break
                        }
                    }
                                           #Document Mask
                    if($word.Text -cmatch 'AS[A-Z]{3}-[0-9]{2}-[0-9]{6}-[0-9]{4}') {
                        $drec = [Doc2Doc]::new()
                        $drec.source_doc_id = $file.BaseName
                        $drec.ref_doc_id = $word.Text -replace '[()]|\/.*|,.*|\..*|\.\|&|\.\|".*|"[A-Z]]{6,}-?\|:' , ''
                        $drec.DATE = $date
                        $docs += $drec
                    }
                }
                # Check for direct instrument tag
                foreach($word in $page_words){
                    if($word.Text -cmatch "^(" + $Inst_Regexes + ")-[0-9]{6}$") {
                        $record = [Tag2Doc]::new()
                        $record.Tag_number = "$Tag_number_Prefix-$($word.Text)"
                        $record.Document_number = $file.BaseName
                        $record.ST = "Hand valve custom search"
                        $record.DATE = $date
                        $record.doc_date = $revision_date
                        $record.issue_reason = $reasonText
                        $record.file_full_path = $file.FullName
                        $record.doctype = $doctype
                        $record.doctitle = $doctitle
                        $record.issuance_code = $issuance_code
                        $tags += $record
                        break
                    } 
                }
                if ($doctype -notmatch "PID|PFD|PSD|DID|SLD") {
                    CONTINUE
                }
                # # If the page size isn’t "A1" and the file name doesn’t match a pattern, skip advanced search.
                if($page.Size -ne "A1" -and $file.BaseName -notmatch '[A-Z0-9]{5}-[A-Z0-9]{3,5}-[A-Z]{5}-[0-9]{2}-[A-Z][0-9]{5}-[0-9]{4}' ) {
                    continue
                }



                # Gather potential instrument sequence words
                $inst_sequens_numbers = @()
                foreach($word in $page_words){
                    if($word.Text -match '^[0-9]{6}[A-Z]?$' -or $word.Text -match '^[0-9]{2}[YX0-9]{1,4}[A-D]?$'){
                        $inst_sequens_numbers += $word
                    }
                }

                $offset = 10
                foreach($word in $page_words){
                    if($word.Text -cmatch "^(" + $Inst_Regexes + ")$") {
                        switch($word.TextOrientation) {
                            'Horizontal' {
                                foreach($seq_word in $inst_sequens_numbers){
                                    if( ($word.BoundingBox.Centroid.Y - $offset) -ge $seq_word.BoundingBox.Bottom -and
                                        ($word.BoundingBox.Centroid.Y - $offset) -le $seq_word.BoundingBox.Top -and
                                        $word.BoundingBox.Centroid.X -ge $seq_word.BoundingBox.Left -and
                                        $word.BoundingBox.Centroid.X -le $seq_word.BoundingBox.Right){
                                        $ti = [Tag2Doc]::new()
                                        $ti.Tag_number = "$Tag_number_Prefix-$($word.Text)-$($seq_word.Text)"
                                        $ti.Document_number = $file.BaseName
                                        $ti.ST = "Advanced Instrument Tag Search"
                                        $ti.DATE = $date
                                        $ti.doc_date = $revision_date
                                        $ti.issue_reason = $reasonText
                                        $ti.file_full_path = $file.FullName
                                        $ti.doctype = $doctype
                                        $ti.doctitle = $doctitle
                                        $ti.issuance_code = $issuance_code
                                        $tags += $ti
                                    }
                                }
                            }
                                                     'Rotate270' {
                                foreach($seq_word in $inst_sequens_numbers){
                                    if( ($word.BoundingBox.Centroid.Y) -ge $seq_word.BoundingBox.Bottom -and
                                        ($word.BoundingBox.Centroid.Y) -le $seq_word.BoundingBox.Top -and 
                                        ($word.BoundingBox.Centroid.X + 10) -ge $seq_word.BoundingBox.Left -and 
                                        ($word.BoundingBox.Centroid.X + 10) -le $seq_word.BoundingBox.Right){
                                        $ti = [Tag2Doc]::new()
                                        $ti.Tag_number = "$Tag_number_Prefix-$($word.Text)-$($seq_word.Text)"
                                        $ti.Document_number = $file.BaseName
                                        $ti.ST = "Advanced Instrument Tag Search"
                                        $ti.DATE = $date
                                        $ti.doc_date = $revision_date
                                        $ti.issue_reason = $reasonText
                                        $ti.file_full_path = $file.FullName
                                        $ti.doctype = $doctype
                                        $ti.doctitle = $doctitle
                                        $ti.issuance_code = $issuance_code
                                        $tags += $ti
                                    }
                                }
                            }
                            'Rotate90' {
                                foreach($seq_word in $inst_sequens_numbers){
                                    if( ($word.BoundingBox.Centroid.Y) -ge $seq_word.BoundingBox.Bottom -and
                                        ($word.BoundingBox.Centroid.Y) -le $seq_word.BoundingBox.Top -and 
                                        ($word.BoundingBox.Centroid.X + 10) -ge $seq_word.BoundingBox.Left -and 
                                        ($word.BoundingBox.Centroid.X + 10) -le $seq_word.BoundingBox.Right){
                                        $ti = [Tag2Doc]::new()
                                        $ti.Tag_number = "$Tag_number_Prefix-$($word.Text)-$($seq_word.Text)"
                                        $ti.Document_number = $file.BaseName
                                        $ti.ST = "Advanced Instrument Tag Search"
                                        $ti.DATE = $date
                                        $ti.doc_date = $revision_date
                                        $ti.issue_reason = $reasonText
                                        $ti.file_full_path = $file.FullName
                                        $ti.doctype = $doctype
                                        $ti.doctitle = $doctitle
                                        $ti.issuance_code = $issuance_code
                                        $tags += $ti
                                    }
                                }
                            }
                            default {
                                $mess = "$($word.Text) - Rotation: $($word.TextOrientation)"
                                Write-Log -Level DEBUG -Message "Cannot capture orientation for: $mess"
                            }
                        }
                    }
                }
            }

                $sw.Stop()
                Write-Log -Level DEBUG -Message ("Processed file '{0}' in {1:N2} seconds." -f $file.Name, $sw.Elapsed.TotalSeconds)
                $filetimes += $sw.Elapsed.TotalSeconds
        }


        return [PSCustomObject]@{
            Tags = $tags | Where-Object {-not ([string]::IsNullOrEmpty($_.Tag_number))}
            Docs = $docs | Where-Object {-not ([string]::IsNullOrEmpty($_.source_doc_id))}
            Times = $filetimes
          }

    } -ArgumentList $temp, $local_path, $epc, $global:DEBUG_ENABLED, $global:scriptname
    $jobs_list += $job

    $offset += $partSize + ([bool] $extraSize)
    if ($extraSize) { --$extraSize }
}

################################################################################
# Wait for all jobs & merge results
################################################################################
Wait-Job -Job $jobs_list
Write-Log -Level INFO -Message "Waiting for all jobs to complete for EPCIC $epc"

$tagexport = @()
$docexport = @()
$alltimes = @()
foreach ($j in $jobs_list) {
    $results = Receive-Job -Job $j
    if($results) {
        $tagexport += $results.Tags
        $docexport += $results.Docs
        $alltimes += $results.Times
    }
}
Remove-Job -Job $jobs_list -Force



$tagexport = $tagexport | Where-Object { -not ([string]::IsNullOrEmpty($_.Document_number)) } | 
          Select-Object Tag_number, Document_number, doctitle, doctype, issuance_code, ST, DATE, doc_date, issue_reason, file_full_path
$docexport = $docexport |  Where-Object { -not ([string]::IsNullOrEmpty($_.ref_doc_id)) }
            Select-Object source_doc_id, ref_doc_id, DATE
$docCount = ($docexport | Where-Object {-not([string]::IsNullOrEmpty($_.ref_doc_id))}).Count
       
Write-Log -Level INFO -Message "Exporting $($tagexport.Count) TAG records to $tag_report"
try {
    $tagexport | Export-Csv -Path $tag_report -NoTypeInformation -Encoding UTF8 -Force
    Write-Log -Level INFO -Message "TAG Export finished successfully."
}
catch {
    Write-Log -Level ERROR -Message "Failed to export CSV to $tag_report. Error: $($_.Exception.Message)"
}

Write-Log -Level INFO -Message "Exporting $($docCount) DOC records to $doc_report"  

try {
    $docexport | Export-Csv -Path $doc_report -NoTypeInformation -Encoding UTF8 -Force
    Write-Log -Level INFO -Message "DOC Export finished successfully."
}
catch {
    Write-Log -Level ERROR -Message "Failed to export CSV to $doc_report. Error: $($_.Exception.Message)"
}

#Average Time elapsed

if($alltimes.Count -gt 0) {
    $avgTime = ($alltimes | Measure-Object -Average).Average
    Write-Log -Level INFO -Message ("Average PDF Processing time: {0:N2} seconds" -f $avgTime)
    $finished = $true
    Write-Log -Level INFO -Message "MultyTread Indexing Completed" -finished $finished
}
else {
    Write-Log -Level WARN -Message "No file processing times recorded."
}

