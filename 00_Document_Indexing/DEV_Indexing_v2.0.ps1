<#
──────────────────────────────────────────────────────────────────────────────
DEV_Indexing_v2.0.ps1 – Optimised multi source tag extractor
Godfather: Mikhail Chestin
Modified by: Daniil Gubin
Last updated : 2025/06/17 , by DG
──────────────────────────────────────────────────────────────────────────────
#>

param (
    [Parameter(Mandatory = $true)]
    [string] $epc,

    [switch] $EnableDebug,
    [switch] $pdf,
    [switch] $excel,
    [switch] $dwg,

    # runspace pool size. Tweak for performance
    [ValidateRange(1,32)]
    [int] $MaxThreads = 10
)

# ─── ENVIRONMENT PREP ── #
if ($EnableDebug) { $Global:DEBUG_ENABLED = $true }


#$root_path = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\Indexing\smallbatch"
$root_path = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing"
if (-not (Test-Path $root_path)) {
    throw "⚠ Root path '$root_path' not found. Check EPC number or network share."
}

$light_regex_config_path = 'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv'


# $logicalProcessors = [Environment]::ProcessorCount
# $scriptCount = 3 # Number of scripts you plan to run in parallel
# $buffer = 2 # Leave a couple of cores for system overhead
# $maxThreads = [Math]::Min(($logicalProcessors - $buffer) / $scriptCount)


$pdf_src_dir = Join-Path $root_path "EPC$epc`_Source"
#$pdf_src_dir = $root_path
$excel_src_dir = Join-Path $root_path "EPC$epc`_Source"
#$excel_src_dir = $root_path
$dwg_src_dir = Join-Path $root_path "EPC$epc`_Source"
#$dwg_src_dir = $root_path

# Output CSVs
$tag_report = Join-Path $root_path "DEV_EPCIC$epc`_PDF_indexing_report.csv"
$excel_output_csv = Join-Path $root_path "DEV_EPCIC$epc`_EXCEL_indexing_report.csv"
$dwg_output_csv = Join-Path $root_path "DEV_EPCIC$epc`_DWG_indexing_report.csv"
$consolidated_csv = Join-Path $root_path "DEV_EPCIC$epc`_indexing_report.csv"

$MasterColumns = @(
    'Tag_number',
    'Document_number',
    'doctitle',
    'doctype',
    'issuance_code',
    'ST',
    'DATE',
    'doc_date',
    'issue_reason',
    'file_full_path',
    'SourceType' # PDF / EXCEL / DWG
)

$python_exe = "C:\ProgramData\anaconda3\python.exe"
$python_script = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\00_Document_Indexing\DEV_xlsx_docs_processing_v3.1.py"

# ─── IMPORT SUPPORT SCRIPTS ────────────────────────────────────────────────

. "$PSScriptRoot\..\Common_Functions.ps1" # logging, PDF helpers
# . "$PSScriptRoot\99.01_Extract_text_from_DWG_func.ps1" # Invoke DwgIndexing. now testing

# ─── SELECTED PIPELINES ───────────────────────────────────────────────────
if (-not ($pdf -or $excel -or $dwg)) { $pdf=$excel=$dwg=$true }

$requested = @(); if ($pdf) { $requested += 'PDF' }
                 if ($excel) { $requested += 'Excel' }
                 if ($dwg) { $requested += 'DWG' }

$requestedcount = $requested.Count

Write-Log -Level INFO -Message "Pipelines requested: $($requested -join ', ')"

Add-Type -AssemblyName System.Collections.Concurrent
$pool = [RunspaceFactory]::CreateRunspacePool(1,$MaxThreads)
$pool.Open()
$tasks = New-Object System.Collections.Concurrent.ConcurrentBag[psobject]

function Start-ParallelJob {
    param([scriptblock]$Script,[hashtable]$arguments=@{})
    $ps = [PowerShell]::Create().AddScript($Script).AddParameters($arguments)
    $ps.RunspacePool = $pool
    $handle = $ps.BeginInvoke()
    $tasks.Add([pscustomobject]@{Handle=$handle;PS=$ps})
}

# ─── PIPELINE LAUNCHERS ───

################################################################################
# --- PDF PROCESSING  ---
################################################################################

if ($pdf) {
    Start-ParallelJob {
        param($src,$out, $root)
        Write-Log -Level INFO -Message "PDF task started…"

        class Tag2Doc {
            [string] $Tag_number
            [string] $Document_number
            [string] $doctitle
            [string] $doctype
            [string] $issuance_code            
            [string] $ST
            [string] $DATE
            [string] $doc_date
            [string] $issue_reason
            [string] $file_full_path
        }
        class Doc2Doc {
            [string] $source_doc_id
            [string] $ref_doc_id
            [string] $DATE
        }

        $date = Get-Date -Format 'MM/dd/yyyy'
        $tags = @()
        $seentags = New-Object System.Collections.Generic.HashSet[string]
        $docs = @()

        . "$using:PSScriptRoot\..\Common_Functions.ps1"

        . "$root\..\Common_Functions.ps1"

        $regexCsvPath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv"
        $Light_Regex = Import-Csv -Delimiter ';' -Path $regexCsvPath

 $Inst_Regexes = @("AAH","AAHH","AI","AIS","AIT","APB","AR","ARC","ASP","AT","AV","BDV","BLD","BPR","BX","BY","CAM","CC","CHV","CI","CMO","CMP","CPF","CPJ","CPR","CS","CTP","CVA","CY","DI","DRS","DT","EJX","EPB","EPR","ESDV","EWS","EX","EY","FA","FAH","FAHH","FAL","FALL","FC","FCS","FCV","FE","FG","FHA","FI","FIT","FIV","FMX","FO","FPS","FQ","FQI","FQV","FQVY","FS","FSH","FSHH","FSL","FSLL","FT","FVI","FVS","FX","FY","GD","GDAH","GDAHH","GDR","GDS","GDT","GLV","GVA","GVAA","HC","HCS","HCV","HD","HDAH","HDAHH","HDC","HDR","HDS","HDT","HF","HG","HGAH","HGAHH","HGS","HIT","HR","HRAH","HS","HSS","HT","HVA","HVS","IAM","ICD","ID","ILK","IMS","IPC","IR","IRAH","JBC","JBE","JBF","JBJ","JBS","LAH","LAHH","LAL","LALL","LC","LCV","LG","LI","LIT","LOS","LRS","LS","LSC","LSD","LSH","LSHH","LSHL","LSL","LSLL","LSS","LT","LVI","LY","MAC","MACA","MCT","MCV","MI","MOV","MRD","MT","MWS","OCP","OWS","PA","PAH","PAHH","PAL","PALL","PB","PC","PCD","PCV","PDAH","PDAHH","PDAL","PDALL","PDC","PDCV","PDI","PDIT","PDRC","PDS","PDSH","PDSHH","PDSL","PDSLL","PDT","PDY","PE","PI","PIT","PRI","PRS","PRV","PS","PSE","PSH","PSHH","PSL","PSLL","PSV","PT","PV","PVI","PX","PY","R","RCU","RD","RO","RTD","RTU","S","SAH","SAHH","SAL","SALL","SCP","SD","SDAH","SDV","SE","SI","SL","SOV","SPR","SS","SSH","SSL","SSSV","ST","SVC","SVP","SWS","SX","SY","TAH","TAHH","TAL","TALL","TC","TCV","TDAH","TDAL","TDIC","TDY","TE","TES","TI","TIT","TMX","TS","TSH","TSHH","TSHL","TSL","TSLL","TSV","TT","TVI","TW","TY","UA","UV","VAH","VAHH","VDU","VGDAH","VGDAHH","VHDAH","VHDAHH","VHGAH","VHGAHH","VHRAH","VIRAH","VMACA","VPSV","VSDAH","VT","WAA","WCV","WMA","WMH","WML","WMR","WMV","WT","X","XA","XAH","XAHH","XC","XCT","XCV","XEP","XI","XL","XPI","XPS","XS","XT","XY","Y","YSL","ZAH","ZAHH","ZE","ZI","ZIC","ZIO","ZL","ZLC","ZLO","ZS","ZSC","ZSO","ZT","2WV","3WV","AAL","ABE","ABI","ABIT","ABT","AC","ACUSH","ACUSL","ADTN","AEN","AGE","AGI","AGIT","AGT","AH","AHA","AHE","AHI","AHIT","AHS","AHT","AIC","AIN","AITN","AME","AMI","AMIT","AMT","AO","AOJ","AP","APCSH","APCSL","ART","ARV","ASCSH","ASCSL","ASH","ASHL","ASL","ASPSH","ASPSL","ATM","ATN","AVY","AWT","AX","AY","AZ","AZE","AZI","AZIT","AZL","AZP","AZR","AZS","AZSC","AZSO","AZT","AZTN","BAL","BDIM","BDIOM","BDOM","BFCL","BFV","BG","BI","BIAD","BIAL","BIALS","BIAS","BIC","BIT","BL","BP","BPV","BR","BRC","BRT","BS","BSG","BSH","BSHL","BSL","BSP","BT","BTC","BTF","BTH","BTHA","BTHL","BTHLI","BTHLR","BTK","BTKL","BTM","BW","BZE","BZL","BZP","BZR","BZS","BZW","CAH","CAL","CE","CGE","CGT","CGTN","CIT","CSC","CSH","CSL","CSO","DE","DHS","DO","DPSH","DPSL","DR","DTT","DX","DY","DZY","EE","EG","EI","EIC","EIT","ER","ERC","ERT","ESD","ESH","ESHL","ESL","ET","EZE","EZI","EZL","EZP","EZR","EZW","FBE","FBI","FBIT","FBS","FBT","FDTN","FEN","FFC","FFE","FFI","FFIC","FFR","FFRC","FFSH","FFSL","FGE","FGI","FGIT","FGN","FGS","FGT","FHE","FHG","FHI","FHIT","FHS","FHT","FICV","FIN","FITN","FITNSH","FITNSL","FJB","FL","FM","FME","FMI","FMIT","FMS","FMT","FQE","FQG","FQIC","FQIT","FQR","FQRC","FQSH","FQSL","FQT","FQX","FQY","FR","FRC","FRT","FSHL","FSV","FTN","FU","FV","FVY","FZ","FZE","FZI","FZIT","FZL","FZP","FZR","FZS","FZSC","FZSO","FZT","FZTN","FZU","FZV","FZW","FZY","FZZS","FZZSC","FZZSO","GAE","GAT","GATN","GAV","GBE","GBT","GBTN","GCE","GCFT","GCFTN","GCT","GCTN","GDE","GDTN","GEE","GET","GETN","GFE","GFT","GFTN","GGAE","GGAT","GGATN","GGBE","GGBT","GGBTN","GGCE","GGCT","GGCTN","GGDE","GGDT","GGDTN","GGE","GGEE","GGET","GGETN","GGFE","GGFT","GGFTN","GGGE","GGGT","GGGTN","GGHE","GGHT","GGHTN","GGIE","GGIT","GGITN","GGJE","GGJT","GGJTN","GGKE","GGKT","GGKTN","GGLE","GGLT","GGLTN","GGME","GGMT","GGMTN","GGNE","GGNT","GGNTN","GGOT","GGOTN","GGPT","GGPTN","GGRT","GGRTN","GGST","GGSTN","GGT","GGTN","GHCT","GHCTN","GHE","GHLT","GHLTN","GHT","GHTN","GIE","GIT","GITN","GJE","GJT","GJTN","GKE","GKT","GKTN","GLE","GLT","GLTN","GMCE","GMCT","GMCTN","GME","GMT","GMTN","GNE","GNOE","GNOT","GNOTN","GNT","GNTN","GOE","GOT","GOTN","GPE","GPOT","GPOTN","GPT","GPTN","GRE","GRT","GRTN","GSE","GST","GSTN","GTT","GTTN","GUE","GUT","GUTN","GVE","GVT","GVTN","GWE","GWT","GWTN","GXE","GXT","GXTN","GYE","GYT","GYTN","GZE","GZT","GZTN","HBB","HBS","HBV","HBY","HHS","HIC","HL","HOA","HSA","HSC","HSO","HSZ","HV","HVY","HX","HY","HYC","HYO","HZIC","HZIO","HZR","HZS","HZSC","HZSO","HZV","HZY","HZZS","HZZSC","HZZSO","IA","II","IIC","IIT","IL","IP","IRC","IRT","ISH","ISHL","ISL","IT","IV","IX","IY","IZE","IZI","IZL","JE","JG","JI","JIC","JIT","JO","JOI","JOR","JQX","JR","JRC","JRT","JSH","JSHL","JSL","JT","JX","JY","JZE","JZI","JZL","JZP","JZR","K","KC","KCV","KE","KG","KI","KIC","KIT","KL","KME","KOG","KOL","KOX","KPE","KQI","KR","KRC","KRT","KSH","KSHL","KSL","KT","KV","KX","KY","KZSC","KZSD","KZSO","KZSS","KZT","KZV","KZY","KZZSC","KZZSO","LA","LACT","LAD","LAHI","LBE","LBI","LBIT","LBP","LBS","LBT","LCI","LCVI","LD","LDI","LDIN","LDIT","LDT","LDTN","LDZT","LDZU","LEI","LEN","LGE","LGI","LGIT","LGS","LGT","LHE","LHG","LHI","LHIT","LHS","LHT","LICI","LIN","LINSH","LINSL","LITN","LITNSH","LITNSL","LL","LLC","LLG","LME","LMI","LMIT","LMS","LMT","LO","LR","LRC","LSHI","LSLI","LTI","LTN","LUX","LV","LVS","LVY","LW","LWG","LX","LZE","LZI","LZIN","LZIT","LZL","LZP","LZS","LZSC","LZSH","LZSL","LZSN","LZSO","LZT","LZTN","LZTNSH","LZTNSL","LZV","LZY","MA","MBOV","MBZSC","MBZSO","MCU","MHHS","MHS","MHT","MLPSH","MLPSL","MZHS","MZLC","MZLO","MZSC","MZSO","MZZS","MZZSC","MZZSO","MZZT","NO","NOC","OM","OV","OWD","OX","PAD","PBI","PBIT","PBS","PBT","PCVN","PDBI","PDBIT","PDBS","PDBT","PDE","PDG","PDGI","PDGIT","PDGN","PDGS","PDGT","PDHCV","PDHG","PDHI","PDHIT","PDHS","PDHSH","PDHSL","PDHT","PDIN","PDITN","PDMI","PDMIT","PDMS","PDMT","PDR","PDRT","PDTN","PDV","PDVY","PDX","PDZI","PDZIT","PDZS","PDZSC","PDZSH","PDZSO","PDZT","PDZTN","PEN","PFL","PFR","PFX","PG","PGI","PGIT","PGN","PGS","PHG","PHI","PHIT","PHS","PHT","PHV","PIN","PITN","PJR","PK","PKL","PKR","PKX","PLC","PLPSH","PLPSL","PMI","PMIT","PMS","PMT","PN","POL","PP","PRC","PRT","PSHL","PSVNSH","PSVNSL","PTC","PTN","PVN","PVY","PZE","PZI","PZIN","PZIT","PZL","PZP","PZS","PZSC","PZSH","PZSL","PZSO","PZT","PZTN","PZTNSH","PZTNSL","PZV","PZY","PZZS","PZZSC","PZZSO","QE","QI","QIC","QIT","QOL","QQ","QQI","QQR","QQX","QR","QRC","QRT","QSH","QSHL","QSL","QT","QY","RAV","RE","REG","RI","RIC","RIT","ROR","ROX","RP","RQ","RQI","RQL","RRC","RRT","RSH","RSHL","RSL","RT","RW","RY","RZE","RZL","RZP","SB","SC","SCN","SCNSH","SCNSL","SCV","SG","SHT","SIC","SIT","SJ","SLO","SME","SMT","SP","SR","SRC","SRT","SRV","SSHL","SSV","SV","SVY","SZE","SZIT","SZT","SZY","TBE","TBI","TBIT","TBS","TBT","TD","TDA","TDC","TDCV","TDE","TDG","TDI","TDIT","TDL","TDR","TDRC","TDRT","TDS","TDSH","TDSL","TDT","TDTN","TDV","TDVY","TDX","TEN","TFI","TFR","TFX","TG","TGE","TGI","TGIT","TGS","TGT","THCV","THE","THG","THI","THIT","THS","THT","THV","TIN","TINSH","TINSL","TIS","TISHL","TITN","TJ","TJE","TJX","TK","TKR","TKX","TL","TME","TMI","TMIT","TMS","TMT","TOR","TP","TR","TRC","TRS","TRT","TSE","TTN","TV","TVY","TZE","TZI","TZIT","TZL","TZP","TZR","TZS","TZSC","TZSH","TZSL","TZSO","TZT","TZTN","TZV","TZW","TZY","UE","UI","UJ","UJR","UL","ULO","UR","USD","UTN","UVY","UX","UY","UZ","UZE","UZL","UZP","UZR","UZS","UZSC","UZSO","UZW","UZY","VBE","VBIT","VBS","VBT","VE","VFD","VG","VGE","VGIT","VGS","VGT","VHE","VHIT","VHS","VHT","VI","VIT","VL","VME","VMIT","VMS","VMT","VP","VR","VRT","VS","VTA","VX","VXE","VXG","VXI","VXL","VXME","VXR","VXT","VXX","VY","VYE","VYME","VYP","VYT","VZ","VZE","VZG","VZIT","VZL","VZP","VZR","VZS","VZT","VZX","WA","WAI","WAR","WAX","WC","WDC","WDCV","WDI","WDIC","WDIT","WDL","WDR","WDRC","WDRT","WDSH","WDSL","WDT","WDX","WE","WEC","WEN","WFI","WFR","WFX","WG","WHS","WI","WIC","WIN","WIT","WKI","WKR","WKX","WL","WMS","WQI","WQL","WQR","WQX","WR","WRT","WS","WSH","WSHL","WSL","WTN","WUL","WUP","WUR","WX","WXE","WXI","WXL","WXR","WXX","WY","WYE","WYL","WZ","WZE","WZIT","WZL","WZR","WZS","WZT","WZX","WZXL","WZXP","WZXR","WZYL","WZYP","WZYR","XBY","XG","XHY","XHZSC","XIO","XLC","XLO","XP","XQ","XR","XV","XVN","XVY","XW","XX","XYC","XYD","XYH","XYL","XYO","XYS","XZC","XZIC","XZIO","XZLC","XZLO","XZO","XZS","XZSC","XZSD","XZSH","XZSL","XZSO","XZSP","XZSS","XZV","XZY","XZYC","XZYO","XZZLC","XZZLO","XZZS","XZZSC","XZZSD","XZZSH","XZZSL","XZZSO","XZZSS","YC","YE","YI","YIC","YL","YR","YSH","YT","YX","YY","YZ","YZE","YZL","YZR","YZX","ZC","ZCV","ZDC","ZDCV","ZDE","ZDG","ZDI","ZDIC","ZDIT","ZDL","ZDR","ZDRC","ZDRT","ZDSH","ZDSL","ZDT","ZDX","ZDXE","ZDXG","ZDXI","ZDXL","ZDXR","ZDXX","ZDY","ZDYE","ZDYG","ZDYI","ZDYL","ZDYR","ZDYX","ZDZE","ZDZI","ZDZL","ZDZR","ZDZX","ZG","ZHS","ZHSC","ZHSO","ZHT","ZIT","ZME","ZMT","ZO","ZP","ZR","ZRC","ZRT","ZSD","ZSE","ZSH","ZSHL","ZSL","ZSP","ZST","ZTN","ZUI","ZUL","ZUP","ZUR","ZW","ZX","ZXE","ZXG","ZXI","ZXL","ZXR","ZY","ZYE","ZYG","ZYL","ZYR","ZYX","ZZ","ZZI","ZZX","ZZXE","ZZXI","ZZXL","ZZXP","ZZXR","ZZYE","ZZYI","ZZYL","ZZYP","ZZYR") -join "|"

$lightregexcompiled = foreach ($row in $light_regex) {

    #adding optional comma or semicolon before end-of-string

    $pattern = $row.Regexp -replace '\$$','(,|;)?$'

    [regex]::new($pattern,
    'Compiled,IgnoreCase')    

    # [regex]::new($row.Regexp,
    # 'Compiled,IgnoreCase')
}

$instregexsequence = [regex]::new("^($Inst_Regexes)-[0-9]{6}$",'Compiled,IgnoreCase')

$InstRegexPrefix = [regex]::new("^($Inst_Regexes)$", 'Compiled,IgnoreCase')

 foreach ($dll in @(
    'UglyToad.PdfPig.dll',
    'UglyToad.PdfPig.DocumentLayoutAnalysis.dll',
    'BouncyCastle.Crypto.dll',
    'itextsharp.dll',
    'itextsharp.pdfa.dll'
)) {
    $path = Join-Path $root "lib\$dll"
    if (Test-Path $path) {
        try { Add-Type -Path $path -ErrorAction SilentlyContinue } catch {}
    }
}

$files = Get-ChildItem -Path $src -Recurse -Include *.pdf -File
Write-Log -Level INFO -Message "PDF task: found $($files.Count) files."

foreach ($file in $files) {

    $meta = $file.FullName.Replace('.pdf', '_null.xml')
    if (-not (Test-Path $meta)) { continue }

    try { [xml]$Xml = Get-Content $meta }
    catch { continue }

    $ns = @{ a = 'http://www.aveva.com/VNET/eiwm' }
    $revision_date = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_revision_date']/a:Value" -Namespace $ns).Node.'#text' -split ' ' | Select-Object -First 1
    $doctype = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_doc_type']/a:Value" -Namespace $ns).Node.InnerText
    $doctitle = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='title']/a:Value" -Namespace $ns).Node.InnerText
    $issuance_code = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_last_return_code']/a:Value" -Namespace $ns).Node.InnerText
    $reasonText = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_revision_object']/a:Value" -Namespace $ns).Node.'#text'

    if ($reasonText -match 'CLD') { continue }

    #  Bookmarks via iTextSharp #
    try {
        $reader = [iTextSharp.text.pdf.PdfReader]::new($file.FullName)
        $bookmarks = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($reader)
        if ($bookmarks) {
            Get-Bookmarks -Bookmarks $bookmarks `
                          -Light_Regex $Light_Regex `
                          -fileBaseName $file.BaseName `
                          -date $date `
                          -revision_date $revision_date `
                          -reasonText $reasonText `
                          -fileFullName $file.FullName `
                          -doctype $doctype `
                          -doctitle $doctitle `
                          -issuance_code $issuance_code `
                          -tags ([ref]$tags)
        }
    } catch {}

    # ---------- 3. Full text search via PdfPig ---------------------
    try { $pdf = [UglyToad.PdfPig.PdfDocument]::Open($file.FullName) }
    catch { continue }

    $nameParts = $file.Name -split '-'
    $tagPrefix = if ($nameParts.Length -ge 3) { $nameParts[2].Substring(0, 4) + 'A' } else { 'XXXXA' }

    for ($p = 1; $p -le $pdf.NumberOfPages; $p++) {
        $page = $pdf.GetPage($p)
        if (-not $page) { continue }

        $page_words = $page.GetWords([UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor]::Instance)
        if (-not $page_words) { continue }

        foreach ($word in $page_words) {

            if([string]::IsNullOrEmpty($word.Text)) { continue }
        
            #NEW - split a long word into potential tag tokens

            $tokens = $word.Text -split '\s*[,;]\s*' | Where-Object {$_}

            foreach ($text in $tokens) {

            if ($text.Text.Length -gt 60) { continue }  #optional santity

            $dupKey = "$text|$file.BaseName|PDF"
            if ($seenTags.Contains($dupKey)) { continue }

            $matched = $false
    
            # --- Light Regex ---

            # foreach ($regex in $Light_Regex) {        #OLD            
            #     if ($word.Text -match $regex.Regexp.Replace('$','')) {        #OLD       

            for ($i = 0; $i -lt $LightRegexCompiled.Count; $i++) {              #CAUTION. NEW replacement (improving performance)
                # if ($LightRegexCompiled[$i].IsMatch($word.Text)) {                #CAUTION. NEW replacement (improving performance)
                if ($LightRegexCompiled[$i].IsMatch($text)) {                #CAUTION. NEW replacement (improving performance)
                    $regex = $light_regex[$i]                                    #CAUTION. NEW replacement (improving performance)
                    $t = [Tag2Doc]::new()
                    # $t.Tag_number = $word.Text
                    $t.Tag_number = $text
                    $t.Document_number = $file.BaseName
                    $t.doctitle = $doctitle
                    $t.doctype = $doctype
                    $t.issuance_code = $issuance
                    $t.ST = $regex.Naming_template_ID
                    $t.DATE = $date
                    $t.doc_date = $revision_date
                    $t.issue_reason = $reasonText
                    $t.file_full_path = $file.FullName
                    $tags += $t
                    $seentags.Add($dupKey) | Out-Null
                    $matched = $true
                    break
                }
            }

            if ($matched) {continue}

            # --- direct doc references -----

            if ($word.Text -cmatch 'AS[A-Z]{3}-[0-9]{2}-[0-9]{6}-[0-9]{4}') {
                $d = [Doc2Doc]::new()
                $d.source_doc_id = $file.BaseName
                $d.ref_doc_id = $w.Text -replace '[()]|\/.*|,.*|\..*'
                $d.DATE = $date
                $docs += $d
            }

            # --- instrument tags (prefix inst seq) ----

            # if ($w.Text -cmatch "^($Inst_Regexes)-[0-9]{6}$") {       #OLD
            # if ($InstRegexSequence.IsMatch($word.Text)) {               #CAUTION. NEW replacement (improving performance)
            if ($InstRegexSequence.IsMatch($text)) {               #CAUTION. NEW replacement (improving performance)
                if ($seentags.Add($dupKey)) {
                $t = [Tag2Doc]::new()
                # $t.Tag_number = "$tagPrefix-$($word.Text)"
                $t.Tag_number = "$tagPrefix-$text"
                $t.Document_number = $file.BaseName
                $t.doctitle = $doctitle
                $t.doctype = $doctype
                $t.issuance_code = $issuance
                $t.ST = 'Hand valve custom search'
                $t.DATE = $date
                $t.doc_date = $revision_date
                $t.issue_reason = $reasonText
                $t.file_full_path = $file.FullName
                $tags += $t

                }
            }
        }
    }

         # 4) ── Advanced Instrument Tag Search (prefix + neighbouring seq) ────

    if ($doctype -notmatch 'PID|PFD|PSD|DID|SLD') { continue }   # skip non P&ID

    if ($page.Size -ne "A1" -and $file.BaseName -notmatch '[A-Z0-9]{5}-[A-Z0-9]{3,5}-[A-Z]{5}-[0-9]{2}-[A-Z][0-9]{5}-[0-9]{4}') {continue}

    # collect sequence number candidates once
    $seqWords = $page_words | Where-Object {
        $_.Text -match '^[0-9]{6}[A-Z]?$' -or $_.Text -match '^[0-9]{2}[A-Z0-9]{1,4}[A-D]?$'
    }

    # $seqWords = @()
    # foreach ($word in $page_words) {
    #     if($word.Text -match '^[0-9]{6}[A-Z]?$' -or $word.Text -match '^[0-9]{2}[YX0-9]{1,4}[A-D]?$') {
    #         $seqWords += $word
    #     }
    # }

    $offset = 10

    foreach ($word in $page_words) {

        if (-not $InstRegexPrefix.IsMatch($word.Text)) {continue}

            switch ($word.TextOrientation) {

                'Horizontal' {
                    foreach ($seq in $seqWords) {
                        if ( ($word.BoundingBox.Centroid.Y - $offset) -ge $seq.BoundingBox.Bottom -and
                             ($word.BoundingBox.Centroid.Y - $offset) -le $seq.BoundingBox.Top   -and
                             $word.BoundingBox.Centroid.X             -ge $seq.BoundingBox.Left  -and
                             $word.BoundingBox.Centroid.X             -le $seq.BoundingBox.Right) {
    
                            $ti                    = [Tag2Doc]::new()
                            $ti.Tag_number         = "$tagPrefix-$($word.Text)-$($seq.Text)"
                            $ti.Document_number    = $file.BaseName
                            $ti.doctitle           = $doctitle
                            $ti.doctype            = $doctype
                            $ti.issuance_code      = $issuance
                            $ti.ST                 = 'Advanced Instrument Tag Search'
                            $ti.DATE               = $date
                            $ti.doc_date           = $revision_date
                            $ti.issue_reason       = $reasonText
                            $ti.file_full_path     = $file.FullName
                            $tags                += $ti
                        }
                    }
                }
    
                'Rotate270' {
                    foreach ($seq in $seqWords) {
                        if ( ($word.BoundingBox.Centroid.Y)          -ge $seq.BoundingBox.Bottom -and
                             ($word.BoundingBox.Centroid.Y)          -le $seq.BoundingBox.Top    -and
                             ($word.BoundingBox.Centroid.X + 10)     -ge $seq.BoundingBox.Left   -and
                             ($word.BoundingBox.Centroid.X + 10)     -le $seq.BoundingBox.Right) {
    
                            $ti                    = [Tag2Doc]::new()
                            $ti.Tag_number         = "$tagPrefix-$($word.Text)-$($seq.Text)"
                            $ti.Document_number    = $file.BaseName
                            $ti.doctitle           = $doctitle
                            $ti.doctype            = $doctype
                            $ti.issuance_code      = $issuance
                            $ti.ST                 = 'Advanced Instrument Tag Search'
                            $ti.DATE               = $date
                            $ti.doc_date           = $revision_date
                            $ti.issue_reason       = $reasonText
                            $ti.file_full_path     = $file.FullName
                            $tags                += $ti
                        }
                    }
                }

                'Rotate90' {
                    foreach ($seq in $seqWords) {
                        if ( ($word.BoundingBox.Centroid.Y)          -ge $seq.BoundingBox.Bottom -and
                             ($word.BoundingBox.Centroid.Y)          -le $seq.BoundingBox.Top    -and
                             ($word.BoundingBox.Centroid.X + 10)     -ge $seq.BoundingBox.Left   -and
                             ($word.BoundingBox.Centroid.X + 10)     -le $seq.BoundingBox.Right) {
    
                            $ti                    = [Tag2Doc]::new()
                            $ti.Tag_number         = "$tagPrefix-$($word.Text)-$($seq.Text)"
                            $ti.Document_number    = $file.BaseName
                            $ti.doctitle           = $doctitle
                            $ti.doctype            = $doctype
                            $ti.issuance_code      = $issuance
                            $ti.ST                 = 'Advanced Instrument Tag Search'
                            $ti.DATE               = $date
                            $ti.doc_date           = $revision_date
                            $ti.issue_reason       = $reasonText
                            $ti.file_full_path     = $file.FullName
                            $tags                += $ti
                        }
                    }
                }

                default {
                    Write-Log -Level DEBUG -Message "Unhandled orientation $($word.TextOrientation) on page $p in $($file.Name)"
                }  
            }   
        }
    }
}

# ---------- 4. Export CSV for this runspace ------------------------
# $all = $tags + $docs
$all = $tags
if ($all.Count) {
    $all | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8 -Force
    Write-Log -Level INFO -Message "PDF task wrote $($all.Count) rows → $out"
} else {
    Write-Log -Level WARN -Message "PDF task found no matches to export."
}
    
    } @{ src=$pdf_src_dir;
        out=$tag_report;
        root = $PSScriptRoot
         }
}

################################################################################
# --- EXCEL EXTRACTOR  ---
################################################################################

if ($excel) {
    Start-ParallelJob {
        param($src,$out,$exe,$script)
        Write-Log -Level INFO -Message "Excel (Python) task started…"
        $env:FOLDER_PATH = $src
        $env:OUTPUT_CSV_PATH = $out
        $env:REGEX_CONFIG = $reg
        & $exe $script
        if ($LASTEXITCODE -ne 0) {
            throw "Excel pipeline failed (exit $LASTEXITCODE)"
        }
    } @{ src=$excel_src_dir;
         out=$excel_output_csv;
         reg=$light_regex_config_path;
         exe=$python_exe;
         script=$python_script }
}

################################################################################
# --- DWG EXTRACTOR  ---
################################################################################

# if ($dwg) {
#     Start-ParallelJob {
#         param($src,$out)
#         Write-Log -Level INFO -Message "DWG task started…"
#         Invoke-DwgIndexing -SourceDir $src -OutCsv $out
#     } @{ src=$dwg_src_dir; out=$dwg_output_csv }
# }

# ─── WAIT FOR ALL RUNSPACES ──

foreach ($t in $tasks) {
    $t.PS.EndInvoke($t.Handle)
    $t.PS.Dispose()
}
$pool.Close(); $pool.Dispose()

Write-Log -Level INFO -Message "All pipelines finished."


if ($requestedCount -gt 1) {


    ###   NORMALIZED MERGING (DOCs are coming soon) ###

    function Merge-CsvReports {
        param (
            [string] $Target,
            [string[]] $CsvList,
            [string[]] $Schema
        )
    
        $existing = $CsvList | Where-Object { Test-Path $_ }
        if ($existing.Count -lt 2) { return $false }
    
        Write-Log -Level INFO -Message "Building consolidated report…"
    
        # 1) Read all rows and tag with their source ---------------------------
        $rows = foreach ($csv in $existing) {
            $srcTag = ([IO.Path]::GetFileNameWithoutExtension($csv)) `
                      -replace '.*_(PDF|EXCEL|DWG).*', '$1'
            Import-Csv $csv | Add-Member NoteProperty SourceType $srcTag -PassThru
        }
    
        if (-not $rows) {
            Write-Log -Level WARN -Message "No data found in partial CSVs."
            return $false
        }
    
        # 2) Normalise every row to the fixed schema ---------------------------
        $normalised = foreach ($r in $rows) {
            $h = [ordered]@{}
            foreach ($col in $Schema) { $h[$col] = $null } # pre seed
    
            foreach ($p in $r.PSObject.Properties) {
                if ($h.Contains($p.Name)) { $h[$p.Name] = $p.Value }
            }
    
            if (-not $h.RecordType) {
                $h.RecordType = if ($h.Tag_number) { 'TAG' } else { 'DOC' }
            }
    
            [pscustomobject]$h
        }
    
        # 3) De duplicate **within each pipeline**
        $deduped = @{ }
        foreach ($r in $normalised) {
            $key = if ($r.RecordType -eq 'TAG') {
                "{0}|{1}|{2}" -f $r.Tag_number,$r.Document_number,$r.SourceType, $r.ST
            } else {
                "{0}|{1}|{2}" -f $r.source_doc_id,$r.ref_doc_id,$r.SourceType
            }
            if (-not $deduped.ContainsKey($key)) { $deduped[$key] = $r }
        }
    
        # 4) Ensure destination folder exists 
        New-Item -Path (Split-Path $Target) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    
        # 5) Export in canonical order
        $deduped.Values |
            Select-Object $Schema |
            Export-Csv -Path $Target -NoTypeInformation -Encoding UTF8 -Force
    
        Write-Log -Level INFO -Message "Consolidated in $Target"
        return $true
    }
    


    #WITHOUT De-dupe and master column START##

    # function Merge-CsvReports {
    #     param (
    #         [string] $Target,
    #         [string[]] $CsvList
    #     )

    #     $existing = $CsvList | Where-Object { Test-Path $_ }
    #     if ($existing.Count -lt 2) { return $false } # nothing to merge

    #     Write-Log -Level INFO -Message "Building consolidated report…"

    #     $merged = foreach ($csv in $existing) {
    #         $tag = ([IO.Path]::GetFileNameWithoutExtension($csv)) `
    #                -replace '.*_(PDF|EXCEL|DWG).*', '$1'
    #         Import-Csv $csv | Add-Member NoteProperty SourceType $tag -PassThru
    #     }

    #     # Ensure target folder exists
    #     New-Item -Path (Split-Path $Target) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

    #     $merged | Export-Csv -Path $Target -NoTypeInformation -Encoding UTF8 -Force
    #     Write-Log -Level INFO -Message "Consolidated → $Target"
    #     return $true
    # }

    #WITHOUT De-dupe and master column END##



    ##############################################################
    # Preserve ALL of the columns regardless START (non-ordered)
    ################################################################

    # function Merge-CsvReports {
    #     param (
    #         [string] $Target,
    #         [string[]] $CsvList
    #     )
    
    #     $existing = $CsvList | Where-Object { Test-Path $_ }
    #     if ($existing.Count -lt 2) { return $false }
    
    #     Write-Log -Level INFO -Message "Building consolidated report…"
    
    #     # ── 1) Slurp all rows into memory
    #     $rows = foreach ($csv in $existing) {
    #         $tag = ([IO.Path]::GetFileNameWithoutExtension($csv)) `
    #                -replace '.*_(PDF|EXCEL|DWG).*', '$1'
    #         Import-Csv $csv | Add-Member NoteProperty SourceType $tag -PassThru
    #     }
    
    #     if (-not $rows) { return $false }
    
    #     # ── 2) Build a union of *all* property names found 
    #     $allProps = $rows |
    #                 ForEach-Object { $_.PSObject.Properties.Name } |
    #                 Sort-Object -Unique
    
    #     # ── 3) Ensure every row has every property (missing → $null)
    #     $rows | ForEach-Object {
    #         foreach ($p in $allProps) {
    #             if (-not $_.PSObject.Properties[$p]) {
    #                 $_ | Add-Member NoteProperty $p $null
    #             }
    #         }
    #     }
    
    #     # ── 4) Export in a predictable column order
    #     $rows | Select-Object $allProps |
    #            Export-Csv -Path $Target -NoTypeInformation -Encoding UTF8 -Force
    
    #     Write-Log -Level INFO -Message "Consolidated → $Target"
    #     return $true
    # }

    #######################################
    # Preserve ALL of the columns regardless END
    #######################################
    
    $mergedOk = Merge-CsvReports -Target $consolidated_csv `
                                 -CsvList @($tag_report, $excel_output_csv, $dwg_output_csv)`
                                -Schema $MasterColumns #Only needed if normalized consolidation used

    # Delete individual files ONLY if merge succeeded
    if ($mergedOk) {
        foreach ($partial in @($tag_report, $excel_output_csv, $dwg_output_csv)) {
            if (Test-Path $partial) {
                Remove-Item $partial -Force -ErrorAction SilentlyContinue
                Write-Log -Level INFO -Message "Removed temporary file $partial"
            }
        }
    }
}
else {
    Write-Log -Level INFO -Message "Single pipeline - keeping individual report only."
}


 
                