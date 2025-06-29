<#
──────────────────────────────────────────────────────────────────────────────
DEV_Indexing_v2.2.ps1 – Optimised multi source tag extractor
Godfather: Mikhail Chestin
Modified by: Daniil Gubin
──────────────────────────────────────────────────────────────────────────────
#>

param(
    [Parameter(Mandatory)][string]$epc,
    [switch]$EnableDebug, [switch]$pdf, [switch]$excel, [switch]$dwg, [switch]$merge,
    [ValidateRange(1, 64)][int]$MaxThreads
)

# ── Environment -----------------------------------------------------------
if ($EnableDebug) { $Global:DEBUG_ENABLED = $true }

$logicalCPU = [Environment]::ProcessorCount
# if($env:INDEXER_MAX_CPU -as [int]){ $logicalCPU=[int]$env:INDEXER_MAX_CPU }

#EXPERIMENT
$logicalCPU = [int]($logicalCPU * 1.5)
if (-not $PSBoundParameters.ContainsKey('MaxThreads')) { $MaxThreads = $logicalCPU }


$commonpath = Resolve-Path "$PSScriptRoot\..\Common_Functions.ps1"

. "$commonpath"

$root_path = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing"
#$root_path = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\Indexing\smallbatch"
if (-not(Test-Path $root_path)) { throw "Root path $root_path not found" }

$light_regex_config_path = 'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv'

$pdf_src_dir = Join-Path $root_path "EPC$epc`_Source"
#$pdf_src_dir = $root_path
$excel_src_dir = Join-Path $root_path "EPC$epc`_Source"
#$excel_src_dir = $root_path
$dwg_src_dir = Join-Path $root_path "EPC$epc`_Source"
#$dwg_src_dir = $root_path

$tag_report = Join-Path $root_path "DEV_EPCIC$epc`_PDF_indexing_report.csv"
$doc_report = Join-Path $root_path "DEV_EPCIC$epc`_doc_PDF_indexing_report.csv"
$excel_output_csv = Join-Path $root_path "DEV_EPCIC$epc`_EXCEL_indexing_report.csv"
$dwg_output_csv = Join-Path $root_path "DEV_EPCIC$epc`_DWG_indexing_report.csv"
$consolidated_csv = Join-Path $root_path "DEV_EPCIC$epc`_indexing_report.csv"

$MasterColumns = @('Tag_number', 'Document_number', 'doctitle', 'doctype', 'issuance_code',
    'ST', 'DATE', 'doc_date', 'revision', 'issue_reason', 'file_full_path', 'SourceType')

$python_exe = "C:\ProgramData\anaconda3\python.exe"
$python_script = "C:\Users\DGU103\Downloads\GIT\DBS-PS_Aveva\00_source_data_processing\00_Document_Indexing\DEV_xlsx_docs_processing_v3.1.py"

# ── Request flags ---------------------------------------------------------
# if (-not($pdf -or $excel -or $dwg)) { $pdf = $excel = $dwg = $true }
$requested = @(); if ($pdf) { $requested += 'PDF' }; if ($excel) { $requested += 'Excel' }; if ($dwg) { $requested += 'DWG' }

Write-Log INFO "Pipelines requested: $($requested -join ', ')"

# ── Runspace pool ---------------------------------------------------------
Add-Type -AssemblyName System.Collections.Concurrent
$pool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads)
$pool.Open()
$tasks = New-Object System.Collections.Concurrent.ConcurrentBag[psobject]

function Start-ParallelJob {
    param([scriptblock]$Script, [hashtable]$arguments = @{})
    $ps = [PowerShell]::Create().AddScript($Script).AddParameters($arguments)
    $ps.RunspacePool = $pool
    $tasks.Add([pscustomobject]@{Handle = $ps.BeginInvoke(); PS = $ps })
}

function Split-Batch {
    param([Parameter(ValueFromPipeline)][object]$Item, [int]$Size)
    begin { $bucket = @() } process {
        $bucket += $Item
        if ($bucket.Count -ge $Size) { , $bucket; $bucket = @() }
    } end { if ($bucket) { , $bucket } }
}

############################################################################
# PDF EXTRACTOR – parallel batches
############################################################################
if ($pdf) {
    $allPdf = Get-ChildItem -Path $pdf_src_dir -Recurse -Include *.pdf -File
    Write-Log INFO "PDF task: found $($allPdf.Count) file(s)."
    $chunkSize = [Math]::Max(240, [Math]::Ceiling($allPdf.Count / $MaxThreads / 2))

    $batchNo = 0

    foreach ($batch in ($allPdf | Split-Batch -Size $chunkSize)) {
        $batchNo++
        $partialCsv = "$tag_report.$batchNo.part"
        $partialCsv_doc = "$doc_report.$batchNo.part"

        Start-ParallelJob {
            param($files, $out, $docsout, $root, $regexCfg, $common)

        #Load common file
            . "$common"

            Write-Log INFO "PDF sub task: processing $($files.Count) files …"

            class Tag2Doc {
                [string]$Tag_number; [string]$Document_number; [string]$doctitle;
                [string]$doctype; [string]$issuance_code; [string]$ST;
                [string]$DATE; [string]$doc_date; [string]$issue_reason;
                [string]$revision; [string]$file_full_path
            }
         
            class Doc2Doc {
                [string] $source_doc_id
                [string] $ref_doc_id
                [string] $revision
                [string] $DATE
            }            

            $date = Get-Date -Format 'MM/dd/yyyy'
            $tags = @()
            $docs = @()

            $processtags = $false
            $processdocs = $true
            $seentags = New-Object System.Collections.Generic.HashSet[string]
            $seendocs = New-Object System.Collections.Generic.HashSet[string]

            $Light_Regex = Import-Csv -Delimiter ';' -Path $regexCfg

            $LightCompiled = $Light_Regex | ForEach-Object {
                [regex]::new(($_.Regexp -replace '\$$', '(,|;)?$'), 'Compiled,IgnoreCase') }

            $Inst_Regexes = @("AAH", "AAHH", "AI", "AIS", "AIT", "APB", "AR", "ARC", "ASP", "AT", "AV", "BDV", "BLD", "BPR", "BX", "BY", "CAM", "CC", "CHV", "CI", "CMO", "CMP", "CPF", "CPJ", "CPR", "CS", "CTP", "CVA", "CY", "DI", "DRS", "DT", "EJX", "EPB", "EPR", "ESDV", "EWS", "EX", "EY", "FA", "FAH", "FAHH", "FAL", "FALL", "FC", "FCS", "FCV", "FE", "FG", "FHA", "FI", "FIT", "FIV", "FMX", "FO", "FPS", "FQ", "FQI", "FQV", "FQVY", "FS", "FSH", "FSHH", "FSL", "FSLL", "FT", "FVI", "FVS", "FX", "FY", "GD", "GDAH", "GDAHH", "GDR", "GDS", "GDT", "GLV", "GVA", "GVAA", "HC", "HCS", "HCV", "HD", "HDAH", "HDAHH", "HDC", "HDR", "HDS", "HDT", "HF", "HG", "HGAH", "HGAHH", "HGS", "HIT", "HR", "HRAH", "HS", "HSS", "HT", "HVA", "HVS", "IAM", "ICD", "ID", "ILK", "IMS", "IPC", "IR", "IRAH", "JBC", "JBE", "JBF", "JBJ", "JBS", "LAH", "LAHH", "LAL", "LALL", "LC", "LCV", "LG", "LI", "LIT", "LOS", "LRS", "LS", "LSC", "LSD", "LSH", "LSHH", "LSHL", "LSL", "LSLL", "LSS", "LT", "LVI", "LY", "MAC", "MACA", "MCT", "MCV", "MI", "MOV", "MRD", "MT", "MWS", "OCP", "OWS", "PA", "PAH", "PAHH", "PAL", "PALL", "PB", "PC", "PCD", "PCV", "PDAH", "PDAHH", "PDAL", "PDALL", "PDC", "PDCV", "PDI", "PDIT", "PDRC", "PDS", "PDSH", "PDSHH", "PDSL", "PDSLL", "PDT", "PDY", "PE", "PI", "PIT", "PRI", "PRS", "PRV", "PS", "PSE", "PSH", "PSHH", "PSL", "PSLL", "PSV", "PT", "PV", "PVI", "PX", "PY", "R", "RCU", "RD", "RO", "RTD", "RTU", "S", "SAH", "SAHH", "SAL", "SALL", "SCP", "SD", "SDAH", "SDV", "SE", "SI", "SL", "SOV", "SPR", "SS", "SSH", "SSL", "SSSV", "ST", "SVC", "SVP", "SWS", "SX", "SY", "TAH", "TAHH", "TAL", "TALL", "TC", "TCV", "TDAH", "TDAL", "TDIC", "TDY", "TE", "TES", "TI", "TIT", "TMX", "TS", "TSH", "TSHH", "TSHL", "TSL", "TSLL", "TSV", "TT", "TVI", "TW", "TY", "UA", "UV", "VAH", "VAHH", "VDU", "VGDAH", "VGDAHH", "VHDAH", "VHDAHH", "VHGAH", "VHGAHH", "VHRAH", "VIRAH", "VMACA", "VPSV", "VSDAH", "VT", "WAA", "WCV", "WMA", "WMH", "WML", "WMR", "WMV", "WT", "X", "XA", "XAH", "XAHH", "XC", "XCT", "XCV", "XEP", "XI", "XL", "XPI", "XPS", "XS", "XT", "XY", "Y", "YSL", "ZAH", "ZAHH", "ZE", "ZI", "ZIC", "ZIO", "ZL", "ZLC", "ZLO", "ZS", "ZSC", "ZSO", "ZT", "2WV", "3WV", "AAL", "ABE", "ABI", "ABIT", "ABT", "AC", "ACUSH", "ACUSL", "ADTN", "AEN", "AGE", "AGI", "AGIT", "AGT", "AH", "AHA", "AHE", "AHI", "AHIT", "AHS", "AHT", "AIC", "AIN", "AITN", "AME", "AMI", "AMIT", "AMT", "AO", "AOJ", "AP", "APCSH", "APCSL", "ART", "ARV", "ASCSH", "ASCSL", "ASH", "ASHL", "ASL", "ASPSH", "ASPSL", "ATM", "ATN", "AVY", "AWT", "AX", "AY", "AZ", "AZE", "AZI", "AZIT", "AZL", "AZP", "AZR", "AZS", "AZSC", "AZSO", "AZT", "AZTN", "BAL", "BDIM", "BDIOM", "BDOM", "BFCL", "BFV", "BG", "BI", "BIAD", "BIAL", "BIALS", "BIAS", "BIC", "BIT", "BL", "BP", "BPV", "BR", "BRC", "BRT", "BS", "BSG", "BSH", "BSHL", "BSL", "BSP", "BT", "BTC", "BTF", "BTH", "BTHA", "BTHL", "BTHLI", "BTHLR", "BTK", "BTKL", "BTM", "BW", "BZE", "BZL", "BZP", "BZR", "BZS", "BZW", "CAH", "CAL", "CE", "CGE", "CGT", "CGTN", "CIT", "CSC", "CSH", "CSL", "CSO", "DE", "DHS", "DO", "DPSH", "DPSL", "DR", "DTT", "DX", "DY", "DZY", "EE", "EG", "EI", "EIC", "EIT", "ER", "ERC", "ERT", "ESD", "ESH", "ESHL", "ESL", "ET", "EZE", "EZI", "EZL", "EZP", "EZR", "EZW", "FBE", "FBI", "FBIT", "FBS", "FBT", "FDTN", "FEN", "FFC", "FFE", "FFI", "FFIC", "FFR", "FFRC", "FFSH", "FFSL", "FGE", "FGI", "FGIT", "FGN", "FGS", "FGT", "FHE", "FHG", "FHI", "FHIT", "FHS", "FHT", "FICV", "FIN", "FITN", "FITNSH", "FITNSL", "FJB", "FL", "FM", "FME", "FMI", "FMIT", "FMS", "FMT", "FQE", "FQG", "FQIC", "FQIT", "FQR", "FQRC", "FQSH", "FQSL", "FQT", "FQX", "FQY", "FR", "FRC", "FRT", "FSHL", "FSV", "FTN", "FU", "FV", "FVY", "FZ", "FZE", "FZI", "FZIT", "FZL", "FZP", "FZR", "FZS", "FZSC", "FZSO", "FZT", "FZTN", "FZU", "FZV", "FZW", "FZY", "FZZS", "FZZSC", "FZZSO", "GAE", "GAT", "GATN", "GAV", "GBE", "GBT", "GBTN", "GCE", "GCFT", "GCFTN", "GCT", "GCTN", "GDE", "GDTN", "GEE", "GET", "GETN", "GFE", "GFT", "GFTN", "GGAE", "GGAT", "GGATN", "GGBE", "GGBT", "GGBTN", "GGCE", "GGCT", "GGCTN", "GGDE", "GGDT", "GGDTN", "GGE", "GGEE", "GGET", "GGETN", "GGFE", "GGFT", "GGFTN", "GGGE", "GGGT", "GGGTN", "GGHE", "GGHT", "GGHTN", "GGIE", "GGIT", "GGITN", "GGJE", "GGJT", "GGJTN", "GGKE", "GGKT", "GGKTN", "GGLE", "GGLT", "GGLTN", "GGME", "GGMT", "GGMTN", "GGNE", "GGNT", "GGNTN", "GGOT", "GGOTN", "GGPT", "GGPTN", "GGRT", "GGRTN", "GGST", "GGSTN", "GGT", "GGTN", "GHCT", "GHCTN", "GHE", "GHLT", "GHLTN", "GHT", "GHTN", "GIE", "GIT", "GITN", "GJE", "GJT", "GJTN", "GKE", "GKT", "GKTN", "GLE", "GLT", "GLTN", "GMCE", "GMCT", "GMCTN", "GME", "GMT", "GMTN", "GNE", "GNOE", "GNOT", "GNOTN", "GNT", "GNTN", "GOE", "GOT", "GOTN", "GPE", "GPOT", "GPOTN", "GPT", "GPTN", "GRE", "GRT", "GRTN", "GSE", "GST", "GSTN", "GTT", "GTTN", "GUE", "GUT", "GUTN", "GVE", "GVT", "GVTN", "GWE", "GWT", "GWTN", "GXE", "GXT", "GXTN", "GYE", "GYT", "GYTN", "GZE", "GZT", "GZTN", "HBB", "HBS", "HBV", "HBY", "HHS", "HIC", "HL", "HOA", "HSA", "HSC", "HSO", "HSZ", "HV", "HVY", "HX", "HY", "HYC", "HYO", "HZIC", "HZIO", "HZR", "HZS", "HZSC", "HZSO", "HZV", "HZY", "HZZS", "HZZSC", "HZZSO", "IA", "II", "IIC", "IIT", "IL", "IP", "IRC", "IRT", "ISH", "ISHL", "ISL", "IT", "IV", "IX", "IY", "IZE", "IZI", "IZL", "JE", "JG", "JI", "JIC", "JIT", "JO", "JOI", "JOR", "JQX", "JR", "JRC", "JRT", "JSH", "JSHL", "JSL", "JT", "JX", "JY", "JZE", "JZI", "JZL", "JZP", "JZR", "K", "KC", "KCV", "KE", "KG", "KI", "KIC", "KIT", "KL", "KME", "KOG", "KOL", "KOX", "KPE", "KQI", "KR", "KRC", "KRT", "KSH", "KSHL", "KSL", "KT", "KV", "KX", "KY", "KZSC", "KZSD", "KZSO", "KZSS", "KZT", "KZV", "KZY", "KZZSC", "KZZSO", "LA", "LACT", "LAD", "LAHI", "LBE", "LBI", "LBIT", "LBP", "LBS", "LBT", "LCI", "LCVI", "LD", "LDI", "LDIN", "LDIT", "LDT", "LDTN", "LDZT", "LDZU", "LEI", "LEN", "LGE", "LGI", "LGIT", "LGS", "LGT", "LHE", "LHG", "LHI", "LHIT", "LHS", "LHT", "LICI", "LIN", "LINSH", "LINSL", "LITN", "LITNSH", "LITNSL", "LL", "LLC", "LLG", "LME", "LMI", "LMIT", "LMS", "LMT", "LO", "LR", "LRC", "LSHI", "LSLI", "LTI", "LTN", "LUX", "LV", "LVS", "LVY", "LW", "LWG", "LX", "LZE", "LZI", "LZIN", "LZIT", "LZL", "LZP", "LZS", "LZSC", "LZSH", "LZSL", "LZSN", "LZSO", "LZT", "LZTN", "LZTNSH", "LZTNSL", "LZV", "LZY", "MA", "MBOV", "MBZSC", "MBZSO", "MCU", "MHHS", "MHS", "MHT", "MLPSH", "MLPSL", "MZHS", "MZLC", "MZLO", "MZSC", "MZSO", "MZZS", "MZZSC", "MZZSO", "MZZT", "NO", "NOC", "OM", "OV", "OWD", "OX", "PAD", "PBI", "PBIT", "PBS", "PBT", "PCVN", "PDBI", "PDBIT", "PDBS", "PDBT", "PDE", "PDG", "PDGI", "PDGIT", "PDGN", "PDGS", "PDGT", "PDHCV", "PDHG", "PDHI", "PDHIT", "PDHS", "PDHSH", "PDHSL", "PDHT", "PDIN", "PDITN", "PDMI", "PDMIT", "PDMS", "PDMT", "PDR", "PDRT", "PDTN", "PDV", "PDVY", "PDX", "PDZI", "PDZIT", "PDZS", "PDZSC", "PDZSH", "PDZSO", "PDZT", "PDZTN", "PEN", "PFL", "PFR", "PFX", "PG", "PGI", "PGIT", "PGN", "PGS", "PHG", "PHI", "PHIT", "PHS", "PHT", "PHV", "PIN", "PITN", "PJR", "PK", "PKL", "PKR", "PKX", "PLC", "PLPSH", "PLPSL", "PMI", "PMIT", "PMS", "PMT", "PN", "POL", "PP", "PRC", "PRT", "PSHL", "PSVNSH", "PSVNSL", "PTC", "PTN", "PVN", "PVY", "PZE", "PZI", "PZIN", "PZIT", "PZL", "PZP", "PZS", "PZSC", "PZSH", "PZSL", "PZSO", "PZT", "PZTN", "PZTNSH", "PZTNSL", "PZV", "PZY", "PZZS", "PZZSC", "PZZSO", "QE", "QI", "QIC", "QIT", "QOL", "QQ", "QQI", "QQR", "QQX", "QR", "QRC", "QRT", "QSH", "QSHL", "QSL", "QT", "QY", "RAV", "RE", "REG", "RI", "RIC", "RIT", "ROR", "ROX", "RP", "RQ", "RQI", "RQL", "RRC", "RRT", "RSH", "RSHL", "RSL", "RT", "RW", "RY", "RZE", "RZL", "RZP", "SB", "SC", "SCN", "SCNSH", "SCNSL", "SCV", "SG", "SHT", "SIC", "SIT", "SJ", "SLO", "SME", "SMT", "SP", "SR", "SRC", "SRT", "SRV", "SSHL", "SSV", "SV", "SVY", "SZE", "SZIT", "SZT", "SZY", "TBE", "TBI", "TBIT", "TBS", "TBT", "TD", "TDA", "TDC", "TDCV", "TDE", "TDG", "TDI", "TDIT", "TDL", "TDR", "TDRC", "TDRT", "TDS", "TDSH", "TDSL", "TDT", "TDTN", "TDV", "TDVY", "TDX", "TEN", "TFI", "TFR", "TFX", "TG", "TGE", "TGI", "TGIT", "TGS", "TGT", "THCV", "THE", "THG", "THI", "THIT", "THS", "THT", "THV", "TIN", "TINSH", "TINSL", "TIS", "TISHL", "TITN", "TJ", "TJE", "TJX", "TK", "TKR", "TKX", "TL", "TME", "TMI", "TMIT", "TMS", "TMT", "TOR", "TP", "TR", "TRC", "TRS", "TRT", "TSE", "TTN", "TV", "TVY", "TZE", "TZI", "TZIT", "TZL", "TZP", "TZR", "TZS", "TZSC", "TZSH", "TZSL", "TZSO", "TZT", "TZTN", "TZV", "TZW", "TZY", "UE", "UI", "UJ", "UJR", "UL", "ULO", "UR", "USD", "UTN", "UVY", "UX", "UY", "UZ", "UZE", "UZL", "UZP", "UZR", "UZS", "UZSC", "UZSO", "UZW", "UZY", "VBE", "VBIT", "VBS", "VBT", "VE", "VFD", "VG", "VGE", "VGIT", "VGS", "VGT", "VHE", "VHIT", "VHS", "VHT", "VI", "VIT", "VL", "VME", "VMIT", "VMS", "VMT", "VP", "VR", "VRT", "VS", "VTA", "VX", "VXE", "VXG", "VXI", "VXL", "VXME", "VXR", "VXT", "VXX", "VY", "VYE", "VYME", "VYP", "VYT", "VZ", "VZE", "VZG", "VZIT", "VZL", "VZP", "VZR", "VZS", "VZT", "VZX", "WA", "WAI", "WAR", "WAX", "WC", "WDC", "WDCV", "WDI", "WDIC", "WDIT", "WDL", "WDR", "WDRC", "WDRT", "WDSH", "WDSL", "WDT", "WDX", "WE", "WEC", "WEN", "WFI", "WFR", "WFX", "WG", "WHS", "WI", "WIC", "WIN", "WIT", "WKI", "WKR", "WKX", "WL", "WMS", "WQI", "WQL", "WQR", "WQX", "WR", "WRT", "WS", "WSH", "WSHL", "WSL", "WTN", "WUL", "WUP", "WUR", "WX", "WXE", "WXI", "WXL", "WXR", "WXX", "WY", "WYE", "WYL", "WZ", "WZE", "WZIT", "WZL", "WZR", "WZS", "WZT", "WZX", "WZXL", "WZXP", "WZXR", "WZYL", "WZYP", "WZYR", "XBY", "XG", "XHY", "XHZSC", "XIO", "XLC", "XLO", "XP", "XQ", "XR", "XV", "XVN", "XVY", "XW", "XX", "XYC", "XYD", "XYH", "XYL", "XYO", "XYS", "XZC", "XZIC", "XZIO", "XZLC", "XZLO", "XZO", "XZS", "XZSC", "XZSD", "XZSH", "XZSL", "XZSO", "XZSP", "XZSS", "XZV", "XZY", "XZYC", "XZYO", "XZZLC", "XZZLO", "XZZS", "XZZSC", "XZZSD", "XZZSH", "XZZSL", "XZZSO", "XZZSS", "YC", "YE", "YI", "YIC", "YL", "YR", "YSH", "YT", "YX", "YY", "YZ", "YZE", "YZL", "YZR", "YZX", "ZC", "ZCV", "ZDC", "ZDCV", "ZDE", "ZDG", "ZDI", "ZDIC", "ZDIT", "ZDL", "ZDR", "ZDRC", "ZDRT", "ZDSH", "ZDSL", "ZDT", "ZDX", "ZDXE", "ZDXG", "ZDXI", "ZDXL", "ZDXR", "ZDXX", "ZDY", "ZDYE", "ZDYG", "ZDYI", "ZDYL", "ZDYR", "ZDYX", "ZDZE", "ZDZI", "ZDZL", "ZDZR", "ZDZX", "ZG", "ZHS", "ZHSC", "ZHSO", "ZHT", "ZIT", "ZME", "ZMT", "ZO", "ZP", "ZR", "ZRC", "ZRT", "ZSD", "ZSE", "ZSH", "ZSHL", "ZSL", "ZSP", "ZST", "ZTN", "ZUI", "ZUL", "ZUP", "ZUR", "ZW", "ZX", "ZXE", "ZXG", "ZXI", "ZXL", "ZXR", "ZY", "ZYE", "ZYG", "ZYL", "ZYR", "ZYX", "ZZ", "ZZI", "ZZX", "ZZXE", "ZZXI", "ZZXL", "ZZXP", "ZZXR", "ZZYE", "ZZYI", "ZZYL", "ZZYP", "ZZYR") -join "|" 
 
            $instregexsequence = [regex]::new("^($Inst_Regexes)-[0-9]{6}$", 'Compiled,IgnoreCase')

            $InstRegexPrefix = [regex]::new("^($Inst_Regexes)$", 'Compiled,IgnoreCase')

            $docmask = [regex]::new('(?<![A-Z0-9])AS[A-Z]{3}-[A-Z0-9]{2}-[0-9]{6}-[0-9]{4}(?![A-Z0-9])', 'Compiled,IgnoreCase')
            
            # --- DLL load -------------------------------------------------------------
            foreach ($dll in 'UglyToad.PdfPig.dll', 'UglyToad.PdfPig.DocumentLayoutAnalysis.dll',
                'BouncyCastle.Crypto.dll', 'itextsharp.dll', 'itextsharp.pdfa.dll') {
                $p = Join-Path $root "lib\$dll"
                if (Test-Path $p) { try { Add-Type -Path $p -ea 0 }catch {} }
            }

            foreach ($file in $files) {

                $meta = $file.FullName.Replace('.pdf', '_null.xml')
                if (-not (Test-Path $meta)) { continue }

                if ($file.BaseName -match 'CLD|SPD|SPD') {continue} #Skipping to the next file (filtering)

                try { [xml]$Xml = Get-Content $meta }
                catch { continue }

                $ns = @{ a = 'http://www.aveva.com/VNET/eiwm' }
                $revision_date = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_revision_date']/a:Value" -Namespace $ns).Node.'#text' -split ' ' | Select-Object -First 1
                $doctype = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_doc_type']/a:Value" -Namespace $ns).Node.InnerText
                $docname = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='cmis:name']/a:Value" -Namespace $ns).Node.InnerText
                $doctitle = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='title']/a:Value" -Namespace $ns).Node.InnerText
                $issuance_code = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_last_return_code']/a:Value" -Namespace $ns).Node.InnerText
                $reasonText = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_revision_object']/a:Value" -Namespace $ns).Node.'#text'
                $revision = (Select-Xml -Xml $Xml -XPath "//a:Characteristic[a:Name='pjc_revision']/a:Value" -Namespace $ns).Node.'#text'

                if ($reasonText -match 'CLD|SPD|SPD') { continue } #Skipping to the next file (filtering)

                #  Bookmarks via iTextSharp #

                if ($processtags) {

                try {
                    $reader = [iTextSharp.text.pdf.PdfReader]::new($file.FullName)
                    $bookmarks = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($reader)
                    if ($bookmarks) {
                        Get-Bookmarks -Bookmarks $bookmarks `
                            -Light_Regex $Light_Regex `
                            -fileBaseName $docname `
                            -date $date `
                            -revision_date $revision_date `
                            -reasonText $reasonText `
                            -fileFullName $file.FullName `
                            -doctype $doctype `
                            -doctitle $doctitle `
                            -revision $revision `
                            -issuance_code $issuance_code `
                            -tags ([ref]$tags)
                    }
                }
                catch {}
            }

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

                    #PROCESS ONLY DOC2DOC relations START

                    if ($processdocs) {

                    foreach ($word in $page_words) {

                        if ([string]::IsNullOrEmpty($word.Text)) { continue }
        
                        #NEW - split a long word into potential tag tokens

                        $tokens = $word.Text -split '\s*[,;/]\s*' | Where-Object { $_ }

                        foreach ($text in $tokens) {

                            if ($text.Text.Length -gt 60) { continue }  #optional santity

                            foreach ($m in $docmask.matches($text)) {

                                $ref = $m.Value #already clean (no rubbish)
                                
                                $d = [Doc2Doc]::new()
                                $d.source_doc_id = $docname
                                $d.ref_doc_id = $ref
                                $d.revision = $revision
                                $d.DATE = $date

                                #deduplicate inside of the doc-indexer itself
                        
                                $dockey = "{0}|{1}" -f $d.source_doc_id, $d.ref_doc_id
                                if ($seenDocs.Add($docKey)) { $docs += $d }

                                }
                            }
                        }

                    }

                            #OLD START
                            # if ($docmask.IsMatch($text)) {
                            # OLD
                               # $clean = $word.Text -replace '^\(|[\)\]\};,/.].*$', '' -replace '^[^\w]+', '' -replace '[^\w]+$', ''
                                
                                #dedupe inside the doc loop (old)
                                # $dockey = "$($file.BaseName)|$clean"
    
                                # if ($seendocs.Add($dockey)) {
    
                                #     $d = [Doc2Doc]::new()
                                #     $d.source_doc_id = $docname
                                #     $d.ref_doc_id = $clean
                                #     $d.revision = $revision
                                #     $d.DATE = $date
                                #     $docs += $d
                                # }

                            #OLD FINISH
            #               }
            #          }
            #     }
            # }

                    #PROCESS ONLY DOC2DOC relations END

                    if ($processtags) {

                    foreach ($word in $page_words) {

                        if ([string]::IsNullOrEmpty($word.Text)) { continue }
        
                        #NEW - split a long word into potential tag tokens

                        $tokens = $word.Text -split '\s*[,;/]\s*' | Where-Object { $_ }

                        foreach ($text in $tokens) {

                            if ($text.Text.Length -gt 60) { continue }  #optional santity

                            if ($docmask.IsMatch($text)) {
                            
                                $clean = $word.Text -replace '^\(|[\)\]\};,/.].*$', '' -replace '^[^\w]+', '' -replace '[^\w]+$', ''
                                
                                #dedupe inside the doc loop
                                $dockey = "$($file.BaseName)|$clean"
    
                                if ($seendocs.Add($dockey)) {
    
                                    $d = [Doc2Doc]::new()
                                    $d.source_doc_id = $docname
                                    $d.ref_doc_id = $clean
                                    $d.revision = $revision
                                    $d.DATE = $date
                                    $docs += $d
                                }
                            }

                        # if ($seendocs.Count) {continue}

                            $dupKey = "$text|$file.BaseName|PDF"
                            if ($seenTags.Contains($dupKey)) { continue }

                            $matched = $false
    
                            # --- Light Regex ---

                            for ($i = 0; $i -lt $LightCompiled.Count; $i++) {
                                if ($LightCompiled[$i].IsMatch($text)) {
                                    $regex = $light_regex[$i]
                                    $t = [Tag2Doc]::new()
                                    $t.Tag_number = $text
                                    $t.Document_number = $docname
                                    $t.doctitle = $doctitle
                                    $t.doctype = $doctype
                                    $t.issuance_code = $issuance_code
                                    $t.ST = $regex.Naming_template_ID
                                    $t.DATE = $date
                                    $t.doc_date = $revision_date
                                    $t.revision = $revision
                                    $t.issue_reason = $reasonText
                                    $t.file_full_path = $file.FullName
                                    $tags += $t
                                    $seentags.Add($dupKey) | Out-Null
                                    $matched = $true
                                    break
                                }
                            }

                            if ($matched) { continue }

                            # --- instrument tags (prefix inst seq) ----

                            if ($InstRegexSequence.IsMatch($text)) {
                                if ($seentags.Add($dupKey)) {
                                    $t = [Tag2Doc]::new()                                    
                                    $t.Tag_number = "$tagPrefix-$text"
                                    $t.Document_number = $docname
                                    $t.doctitle = $doctitle
                                    $t.doctype = $doctype
                                    $t.issuance_code = $issuance_code
                                    $t.ST = 'Hand valve custom search'
                                    $t.DATE = $date
                                    $t.doc_date = $revision_date
                                    $t.revision = $revision
                                    $t.issue_reason = $reasonText
                                    $t.file_full_path = $file.FullName
                                    $tags += $t

                                }
                            }

                        }

                    }

                    # Advanced Instrument Tag Search (prefix + neighbouring seq) ────

                    if ($doctype -notmatch 'PID|PFD|PSD|DID|SLD') { continue }   # skip non P&ID

                    if ($page.Size -ne "A1" -and $file.BaseName -notmatch '[A-Z0-9]{5}-[A-Z0-9]{3,5}-[A-Z]{5}-[0-9]{2}-[A-Z][0-9]{5}-[0-9]{4}') { continue }

                    # collect sequence number candidates once
                    $seqWords = $page_words | Where-Object {
                        $_.Text -match '^[0-9]{6}[A-Z]?$' -or $_.Text -match '^[0-9]{2}[A-Z0-9]{1,4}[A-D]?$'
                    }
                    $offset = 10

                    foreach ($word in $page_words) {

                        if (-not $InstRegexPrefix.IsMatch($word.Text)) { continue }

                        switch ($word.TextOrientation) {

                            'Horizontal' {
                                foreach ($seq in $seqWords) {
                                    if ( ($word.BoundingBox.Centroid.Y - $offset) -ge $seq.BoundingBox.Bottom -and
                             ($word.BoundingBox.Centroid.Y - $offset) -le $seq.BoundingBox.Top -and
                                        $word.BoundingBox.Centroid.X -ge $seq.BoundingBox.Left -and
                                        $word.BoundingBox.Centroid.X -le $seq.BoundingBox.Right) {
    
                                        $ti = [Tag2Doc]::new()
                                        $ti.Tag_number = "$tagPrefix-$($word.Text)-$($seq.Text)"
                                        $ti.Document_number = $docname
                                        $ti.doctitle = $doctitle
                                        $ti.doctype = $doctype
                                        $ti.issuance_code = $issuance_code
                                        $ti.ST = 'Advanced Instrument Tag Search'
                                        $ti.DATE = $date
                                        $ti.doc_date = $revision_date
                                        $ti.revision = $revision
                                        $ti.issue_reason = $reasonText
                                        $ti.file_full_path = $file.FullName
                                        $tags += $ti
                                    }
                                }
                            }
    
                            'Rotate270' {
                                foreach ($seq in $seqWords) {
                                    if ( ($word.BoundingBox.Centroid.Y) -ge $seq.BoundingBox.Bottom -and
                             ($word.BoundingBox.Centroid.Y) -le $seq.BoundingBox.Top -and
                             ($word.BoundingBox.Centroid.X + 10) -ge $seq.BoundingBox.Left -and
                             ($word.BoundingBox.Centroid.X + 10) -le $seq.BoundingBox.Right) {
    
                                        $ti = [Tag2Doc]::new()
                                        $ti.Tag_number = "$tagPrefix-$($word.Text)-$($seq.Text)"
                                        $ti.Document_number = $docname
                                        $ti.doctitle = $doctitle
                                        $ti.doctype = $doctype
                                        $ti.issuance_code = $issuance_code
                                        $ti.ST = 'Advanced Instrument Tag Search'
                                        $ti.DATE = $date
                                        $ti.doc_date = $revision_date
                                        $ti.revision = $revision
                                        $ti.issue_reason = $reasonText
                                        $ti.file_full_path = $file.FullName
                                        $tags += $ti
                                    }
                                }
                            }

                            'Rotate90' {
                                foreach ($seq in $seqWords) {
                                    if ( ($word.BoundingBox.Centroid.Y) -ge $seq.BoundingBox.Bottom -and
                             ($word.BoundingBox.Centroid.Y) -le $seq.BoundingBox.Top -and
                             ($word.BoundingBox.Centroid.X + 10) -ge $seq.BoundingBox.Left -and
                             ($word.BoundingBox.Centroid.X + 10) -le $seq.BoundingBox.Right) {
    
                                        $ti = [Tag2Doc]::new()
                                        $ti.Tag_number = "$tagPrefix-$($word.Text)-$($seq.Text)"
                                        $ti.Document_number = $docname
                                        $ti.doctitle = $doctitle
                                        $ti.doctype = $doctype
                                        $ti.issuance_code = $issuance_code
                                        $ti.ST = 'Advanced Instrument Tag Search'
                                        $ti.DATE = $date
                                        $ti.doc_date = $revision_date
                                        $ti.revision = $revision
                                        $ti.issue_reason = $reasonText
                                        $ti.file_full_path = $file.FullName
                                        $tags += $ti
                                    }
                                }
                            }

                            default {
                                Write-Log -Level DEBUG -Message "Unhandled orientation $($word.TextOrientation) on page $p in $($file.Name)"
                        }  
                    }   
                }

            } # If processing only tags

            }   #Foreach Page
        } #Foreach File

            # ── DEDUP WITHIN THIS BATCH (keep most specific ST) START ────────────────────
            $priority = @(
                'Naming_template_ID', # light regex rows (the real pattern)
                'From bookmarks', # bookmarks injected tags
                'Advanced Instrument Tag Search', # geometry neighbour match
                'Hand valve custom search' # fallback prefix seq
            )

            $best_t = @{} # key = Tag|Doc value = Tag2Doc object
            # $best_d = @{} # key = Doc|Doc value = Doc2Doc object

            foreach ($t in $tags) {
                $key = "{0}|{1}|{2}" -f $t.Tag_number, $t.Document_number, 'PDF'
                if ($best_t.ContainsKey($key)) {
                    $existing = $best_t[$key]

                    # choose the row whose ST appears **earlier** in $priority
                    if ($priority.IndexOf($t.ST) -lt $priority.IndexOf($existing.ST)) {
                        $best_t[$key] = $t
                    }
                }
                else {
                    $best_t[$key] = $t
                }
            }

            # foreach ($d in $docs) {
            #     $dockey = "{0}|{1}" -f $d.source_doc_id, $d.ref_doc_id
            #     if ($best_d.ContainsKey($key)) {
            #         $existing = $best_d[$key]

            #         # choose the row whose ST appears **earlier** in $priority
            #         if ($priority.IndexOf($d.ST) -lt $priority.IndexOf($existing.ST)) {
            #             $best_d[$key] = $d
            #         }
            #     }
            #     else {
            #         $best_d[$key] = $d
            #     }
            # }           

            $tags = $best_t.Values
            # $docs = $best_d.Values

            # ── DEDUP WITHIN THIS BATCH (keep most specific ST) FINISH ────────────────────


            if ($tags.Count) {
                $tags | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8 -Force
                Write-Log INFO "PDF sub task wrote $($tags.Count) rows : $out"
            }
            else { Write-Log WARN "PDF sub task found no tag matches." }

            if ($docs.Count) {
                $docs | Export-Csv -Path "$docsout" -NoTypeInformation -Encoding UTF8 -Force
                Write-Log INFO "PDF sub task wrote $($docs.Count) rows: $docsout"
            }
            else {Write-Log WARN "PDF sub task found no doc matches." }

        } @{ files = $batch; out = $partialCsv; docsout = $partialCsv_doc; root = $PSScriptRoot; regexCfg = $light_regex_config_path; common = $commonpath }
    }
}

############################################################################
# EXCEL EXTRACTOR
############################################################################
if ($excel) {
    Start-ParallelJob {
        param($src, $out, $exe, $script, $reg)

        Write-Log INFO "Excel (Python) task started…"
        $env:FOLDER_PATH = $src; $env:OUTPUT_CSV_PATH = $out; $env:REGEX_CONFIG = $reg
        & $exe $script; if ($LASTEXITCODE) { throw "Excel pipeline failed (exit $LASTEXITCODE)" }

    } @{ src = $excel_src_dir; out = $excel_output_csv; reg = $light_regex_config_path;
        exe = $python_exe; script = $python_script 
    }
}


############################################################################
# DWG EXTRACTOR - will come soon
############################################################################


# One progress bar which tracks all the handles
$jobTotal = $tasks.Count
$completed = 0
$lastPrint = Get-Date

while ($completed -lt $jobTotal) {

    # count finished jobs
    $completed = ($tasks | Where-Object {$_.PS.InvocationStateInfo.State -match 'Completed|Failed'}).Count

    # update at most once per second
    if ((Get-Date) -gt $lastPrint.AddSeconds(1)) {
        $percent = [math]::Round((100 * $completed / $jobTotal), 2)
        Write-Progress -Activity 'DEV_Indexing pipelines' `
                       -Status "$completed / $jobTotal finished ($percent%)" `
                       -PercentComplete $percent
        $lastPrint = Get-Date
    }

    Start-Sleep -Milliseconds 200
}

Write-Progress -Activity 'DEV_Indexing pipelines' -Completed -Status 'All done'


############################################################################
# Wait & close pool / Consolidating Partial Files
############################################################################
foreach ($t in $tasks) { $t.PS.EndInvoke($t.Handle); $t.PS.Dispose() }
$pool.Close(); $pool.Dispose()
Write-Log INFO "All sub tasks finished."

# ─── Merge Tag PDF partial CSVs (skip duplicate headers) ──────────────────
$partFiles = Get-ChildItem "$tag_report.*.part" | Sort-Object Name
if ($partFiles) {

    # read first part completely
    $first = $partFiles | Select-Object -First 1
    Get-Content $first | Set-Content -Encoding UTF8 -Path $tag_report

    # append remaining parts, dropping headers
    foreach ($pf in ($partFiles | Select-Object -Skip 1)) {
        $lines = Get-Content $pf
        if ($lines.Count -gt 1) {
            $lines | Select-Object -Skip 1 | Add-Content -Path $tag_report
        }
    }

    # clean up
    $partFiles | Remove-Item -Force
    Write-Log INFO "PDF: merged $($partFiles.Count) parts : $tag_report"
}

# ─── Merge Doc PDF partial CSVs (skip duplicate headers) ──────────────────
$partFiles2 = Get-ChildItem "$doc_report.*.part" | Sort-Object Name
if ($partFiles2) {

    # read first part completely
    $first2 = $partFiles2 | Select-Object -First 1
    Get-Content $first2 | Set-Content -Encoding UTF8 -Path $doc_report

    # append remaining parts, dropping headers
    foreach ($pf in ($partFiles2 | Select-Object -Skip 1)) {
        $lines2 = Get-Content $pf
        if ($lines2.Count -gt 1) {
            $lines2 | Select-Object -Skip 1 | Add-Content -Path $doc_report
        }
    }

    # clean up
    $partFiles2 | Remove-Item -Force
    Write-Log INFO "PDF: merged $($partFiles2.Count) parts : $doc_report"
}

if($merge) {

# Consolidate if multiple pipelines
# if ($requested.Count -gt 1) {
    function Merge-CsvReports {
        param([string]$Target, [string[]]$CsvList, [string[]]$Schema)
        $existing = $CsvList | Where-Object { Test-Path $_ }; if ($existing.Count -lt 2) { return }
        $rows = foreach ($csv in $existing) {
            $srcTag = ([IO.Path]::GetFileNameWithoutExtension($csv)) -replace '.*_(PDF|EXCEL|DWG).*', '$1'
            Import-Csv $csv | Add-Member NoteProperty SourceType $srcTag -PassThru
        }
        if (-not $rows) { return }
        $normalised = foreach ($r in $rows) {
            $h = [ordered]@{}
            foreach ($c in $Schema) { $h[$c] = $null }
            foreach ($p in $r.PSObject.Properties) {
                if ($h.Contains($p.Name)) { $h[$p.Name] = $p.Value }
            }
            if (-not $h.RecordType) { $h.RecordType = if ($h.Tag_number) { 'TAG' }else { 'DOC' } }
            [pscustomobject]$h
        }


        # OLD SIMPLE DEDUPE START

        # $deduped=@{}       
        # foreach ($r in $norm) {
        #         if ($r.RecordType -eq 'TAG') {
        #         $key = "{0}|{1}|{2}|{3}" -f $r.Tag_number, $r.Document_number, $r.SourceType, $r.ST
        #      } else {
        #     $key = "{0}|{1}|{2}|{3}" -f $r.source_doc_id, $r.ref_doc_id, $r.SourceType, $r.ST
        #          }
            
        #     if (-not $dedup.ContainsKey($key)) {
        #     $dedup[$key] = $r
        #         }
        #     }

        # OLD SIMPLE DEDUPE FINISH

        # Deduplicate across all pipelines

        $priority = @('Naming_template_ID',
            'From bookmarks'
            'Advanced Instrument Tag Search',
            'Hand valve custom search')

        $dedup = @{}
        foreach ($r in $normalised) {

            if ($r.RecordType -eq 'TAG') {
                $key = '{0}|{1}' -f $r.Tag_number, $r.Document_number
                # $key = '{0}|{1}|{2}' -f $r.Tag_number, $r.Document_number, $r.SourceType
                #$key = '{0}|{1}|{2}|{3}' -f $r.Tag_number, $r.Document_number, $r.SourceType, $r.ST


                #OLD LOGIC before PDF WIN START

                # if ($dedup.ContainsKey($key)) {
                #     $old = $dedup[$key]
                #     if ($priority.IndexOf($r.ST) -lt $priority.IndexOf($old.ST)) {
                #         $dedup[$key] = $r
                #     }
                # } 
                # else { $dedup[$key] = $r }

                #OLD LOGIC before PDF WIN END

                if ($dedup.ContainsKey($key)) {
                    $old = $dedup[$key]
                
                    $keepNew =
                        #Prefer PDF over any other SourceType
                        ($old.SourceType -ne 'PDF' -and $r.SourceType -eq 'PDF') -or
                        #If both are PDF (or both not), prefer higher priority ST
                        ( ($r.SourceType -eq $old.SourceType) -and
                          ($priority.IndexOf($r.ST) -lt $priority.IndexOf($old.ST)) )
                
                    if ($keepNew) { $dedup[$key] = $r }
                }

                else {$dedup[$key] = $r}
                    
            }
            else {
                # DOC rows – keep first one, keyed on SourceType. WILL FOLLOW
                $key = '{0}|{1}|{2}' -f $r.source_doc_id, $r.ref_doc_id, $r.SourceType
                if (-not $dedup.ContainsKey($key)) { $dedup[$key] = $r }

            }
        }
          
        New-Item -Path (Split-Path $Target) -ItemType Directory -Force | Out-Null

        $dedup.Values | Select-Object $Schema | Export-Csv -Path $Target -NoTypeInformation -Encoding UTF8 -Force
        Write-Log INFO "Consolidated : $Target"
        return $true

    }
    
    $merged = Merge-CsvReports -Target $consolidated_csv `
        -CsvList @($tag_report, $excel_output_csv, $dwg_output_csv) `
        -Schema $MasterColumns

    if ($merged) {
        foreach ($tmp in @($tag_report, $excel_output_csv, $dwg_output_csv)) {
            if (Test-Path $tmp) { Remove-Item $tmp -Force; Write-Log INFO "Removed temp $tmp" }
        }
    }
# } # Requested flag
# else {
#     Write-Log INFO "Single pipeline - keeping individual reports."
# }

}
