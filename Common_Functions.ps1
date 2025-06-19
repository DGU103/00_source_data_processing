function Write-Log {
    param(
        #[Parameter(Mandatory=$true)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,
        [switch]$NoTimeStamp,

        [bool]$finished
    )

    $timestamp = if ($NoTimeStamp) { "" } else { "$(Get-Date -Format 'MM-dd-yyyy HH:mm:ss') " }
    $logformat = Get-Date -Format 'MM-dd-yyyy'
    $logLine = "$($timestamp)[$Level] $Message"

    $logpath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Logs\Methods\$method\$logformat"

        if(!(test-path -PathType container $logpath)) {
                New-Item -ItemType Directory -Path $logpath
        }

    $logname = "EPCIC$epc" + '_' + $scriptname + '_' + $logformat + '.log'
    $logFile = "$logpath\$logname"

    if (($Level -eq 'DEBUG') -and (-not $global:DEBUG_ENABLED)) {
        return
    }

    # Print to console
    switch ($Level) {
        "INFO" { Write-Host $logLine -ForegroundColor Green }
        "WARN" { Write-Host $logLine -ForegroundColor Yellow }
        "ERROR" { Write-Host $logLine -ForegroundColor Red }
        "DEBUG" { Write-Host $logLine -ForegroundColor Gray }
    }

        # Append to log file
        Add-Content -Path $logFile -Value $logLine


    if ($finished) {Rename-Item -Path $logFile -NewName "OK_$logname" -ErrorAction SilentlyContinue}

    if ($Level -eq 'ERROR') { Rename-Item -Path $logFile -NewName "ERR_$logname" -ErrorAction SilentlyContinue}

}

function Get-Bookmarks {
    param (
        [array]$Bookmarks,
        [array]$Light_Regex,
        [string]$fileBaseName,
        [string]$date,
        [string]$revision_date,
        [string]$reasonText,
        [string]$fileFullName,
        [string]$doctype,
        [string]$doctitle,
        [string]$issuance_code,
        [ref]$tags
    )

    foreach ($bookmark in $Bookmarks) {
        if ($bookmark.Kids) {
            Get-Bookmarks -Bookmarks $bookmark.Kids -Light_Regex $Light_Regex -fileBaseName $fileBaseName -date $date -revision_date $revision_date -reasonText $reasonText -fileFullName $fileFullName -doctype $doctype -doctitle $doctitle -issuance_code $issuance_code -tags $tags
        } else {
            foreach ($regex in $Light_Regex) {
                if ($bookmark.Title.Split()[0] -match $regex.Regexp) {
                    $record = [Tag2Doc]::new()
                    $record.Tag_number = $bookmark.Title.Split()[0]
                    $record.Document_number = $fileBaseName
                    $record.ST = "From bookmarks"
                    $record.DATE = $date
                    $record.doc_date = $revision_date
                    $record.issue_reason = $reasonText
                    $record.file_full_path = $fileFullName
                    $record.doctype = $doctype
                    $record.doctitle = $doctitle
                    $record.issuance_code = $issuance_code
                    $tags.Value += $record
                    break
                }
            }
        }
    }
}

#Indexation Related Processes
function Invoke-PDFIndexing {

    param(
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory)][string]$Epc,
        [Parameter(Mandatory)][string]$LocalPath
    )

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

    $date = Get-Date -Format 'MM/dd/yyyy'

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
            Tag2Doc = $tags | Where-Object {-not ([string]::IsNullOrEmpty($_.Tag_number))}
            Doc2Doc = $docs | Where-Object {-not ([string]::IsNullOrEmpty($_.source_doc_id))}
            Times = $filetimes
          }
}

function Invoke-ExcelIndexing {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory)][string]$Epc,
        [Parameter(Mandatory)][string]$LocalPath
    )

    if(-not $Files){ return [pscustomobject]@{ Tag2Doc=@(); Doc2Doc=@(); Times=@() } }

    class Tag2Doc {
        [string]$Tag_number
        [string]$Document_number
        [string]$doctitle
        [string]$ST
        [string]$DATE
        [string]$doc_date
        [string]$issue_reason
        [string]$file_full_path
        [string]$doctype
        [string]$issuance_code
    }
    class Doc2Doc {
        [string]$source_doc_id
        [string]$ref_doc_id
        [string]$DATE
    }

    $tag2docOut = [Collections.Generic.List[Tag2Doc]]::new()
    $doc2docOut = [Collections.Generic.List[Doc2Doc]]::new()
    $times = @()

    # PERFORMANCE improvement: open ONE Excel instance per job, not per file
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach($file in $Files){

        $identifier = $file.BaseName.Split('_')[0]

        $metapath = Get-ChildItem -path $file.DirectoryName -Filter "*_null.xml" | Where-Object { $_.BaseName.StartsWith($identifier) } | Select-Object -ExpandProperty FullName -First 1
        
        if (-not (Test-Path $metaPath)) {
            # Write-Log -Level ERROR -Message "Metadata for $($file.FullName) not found. Skipping file."
            continue
        }

        try {
            [xml]$XmlDocument = Get-Content $metaPath
        }
        catch {
            # Write-Log -Level ERROR -Message "Unable to read XML for $($file.FullName): $($_.Exception.Message). Skipping."
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
            # Write-Log -Level WARN -Message "Skipping $($file.Name) because reason is CLD."
            continue
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        try{
            $wb = $excel.Workbooks.Open($file.FullName, 3) # read-only, no links
            # $wb = $excel.Workbooks.Open($file.FullName, 3) # read-only, no links

            $tagPattern = [regex]'^[A-Z0-9]+(?:-[A-Z0-9]+){1,}-\d{3,6}$'

            foreach($ws in $wb.Worksheets){

                # read entire sheet in one COM call
                $data = $ws.UsedRange.Value2
                if (-not $data) { continue }

                $rowMax = $data.GetUpperBound(0) # 1-based
                $colMax = $data.GetUpperBound(1)

     
             # Find the header row by locating the cell with "Equipment No"
                # locate header “Equipment No”
                $headerRow = 0
                $equipCol = 0
                for ($r = 1; $r -le $rowMax; $r++) {
                    for ($c = 1; $c -le $colMax; $c++) {
                        $val = $data[$r,$c]
                        if (-not $val) { continue }
                
                        if ($val.ToString().Trim().Equals('Equipment No', 'InvariantCultureIgnoreCase')) {
                            $headerRow = $r
                            $equipCol = $c
                            break
                        }
                
                    }
                    if ($equipCol) { break }
                }
                if (-not $equipCol) { continue } # sheet skipped

                    for ($r = $headerRow + 1; $r -le $rowMax; $r++) {
                        $raw = $data[$r,$equipCol]
                        if (-not $raw) { continue }
        
                        $tag = $raw.ToString().Trim()

                    if (-not $tagPattern.IsMatch($tag)) { continue } # ← filter junk rows

                
                    # $tag = $match.Value
                    $t2d = [Tag2Doc]::new()
                    $t2d.Tag_number = $tag
                    $t2d.Document_number = $file.BaseName
                    $t2d.ST = 'EXL'
                    $t2d.DATE = $date
                    $t2d.doc_date = $revision_date
                    $t2d.issue_reason = $reasonText
                    $t2d.file_full_path = $file.FullName
                    $t2d.doctype = $doctype
                    $t2d.doctitle = $doctitle
                    $t2d.issuance_code = $issuance_code
                    $tag2docOut.Add($t2d)
                    }
    }
        
            $wb.Close($false)
            [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        }
        catch{
            # Write-Log -Level ERROR -Message "Excel read error $($file.Name): $($_.Exception.Message)"
        }
        $sw.Stop()
        $times += $sw.Elapsed.TotalSeconds
    }

    # full COM cleanup
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    return [pscustomobject]@{
        Tag2Doc = $tag2docOut
        Doc2Doc = $doc2docOut 
        Times = $times
    }
}
         
function Update-SqlTagdoc {

    param (
        [int]$epc,
        [string]$aim_report
    )
#defining a connection string for SQL server

$connString = "Server=QA-SQL-TEST2019; Database=AIM_DEV; Integrated Security=True;"
$connection = New-Object System.Data.SqlClient.SqlConnection $connString
$connection.Open()

# Read CSV into a DataTable
$dataTable = New-Object System.Data.DataTable
# Define DataTable columns to match Tag2Doc table structure
[void]$dataTable.Columns.Add("Tag2DocID", [int])
[void]$dataTable.Columns.Add("EPCIC", [int])
[void]$dataTable.Columns.Add("Reference_ID", [string])
[void]$dataTable.Columns.Add("Document_ID", [string])


Import-Csv $aim_report -Delimiter ';' | ForEach-Object {
    $row = $dataTable.NewRow()
    $row["Tag2DocID"] = "1"
    $row["EPCIC"] = $epc
    $row["Reference_ID"] = $_.Reference_ID
    $row["Document_ID"] = $_.Document_ID
    $dataTable.Rows.Add($row)
}

$bulk = New-Object System.Data.SqlClient.SqlBulkCopy($connection)
$bulk.DestinationTableName = "Tag2Doc"

try {
    $bulk.WriteToServer($dataTable)
    Write-Host "Bulk insert complete. Rows inserted: " $dataTable.Rows.Count
    $2count = $dataTable.Rows.Count
}
catch {
    Write-Host "Bulk insert failed: $($_.Exception.Message)"
}
finally {
    $bulk.Close()
    $connection.Close() 
}

}

function Invoke-Tag2DocUpsert {
    param([System.Collections.IEnumerable]$Batch,
    [String]$epc)

$data = New-Object System.Data.DataTable
# Define DataTable columns to match Tag2Doc table structure
[void]$data.Columns.Add("EPCIC", [int])
[void]$data.Columns.Add("Reference_ID", [string])
[void]$data.Columns.Add("Document_ID", [string])


foreach ($r in $Batch) {
    $row = $data.NewRow()
    $row.EPCIC = $epc
    $row.Reference_ID = $r.Reference_ID
    $row.Document_ID = $r.Document_ID
    $data.Rows.Add($row)
}

# Call the proc with a TVP
$conn = New-Object System.Data.SqlClient.SqlConnection `
          "Server=QA-SQL-TEST2019;Database=AIM_DEV;Integrated Security=SSPI"
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandType = [System.Data.CommandType]::StoredProcedure
$cmd.CommandText = "dbo.usp_Tag2Doc_Load"

$param = $cmd.Parameters.Add("@NewRows",
            [System.Data.SqlDbType]::Structured)
$param.TypeName = "dbo.Tag2DocInput"
$param.Value = $data

$cmd.ExecuteNonQuery()
$conn.Close()

}

function E3D {


    param (
        [bool]$fullrun,
        [bool]$packingvoke,
        [string]$e3d_filters,
        [string]$e3d_links,
        [string]$e3d_tags,
        [string]$e3d_model,     
        [string]$epc  
        )

    if ($e3d_tags -eq 'y') {  
        
        & "$PSScriptRoot\02_E3D\00.01_Export_Tags_from_E3D.ps1"
        & "$PSScriptRoot\02_E3D\02.01_Export_parent_links_from_E3D.ps1"
    
    }
       
    
    $scripts = @("01.00_E3D_Tagged_Item_full_regex_Regex_Filtering.ps1",
    "01.01_E3D_Non-Tagged_Items_Filtering.ps1")

    switch ($e3d_filters) {

        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\02_E3D\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\02_E3D\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\02_E3D\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\02_E3D\$script" -epc 11
            & "$PSScriptRoot\02_E3D\$script" -epc 12
            & "$PSScriptRoot\02_E3D\$script" -epc 13   

            }

    }

    }

    switch ($e3d_links) {

        "11" { & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 11
            & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 12
            & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 13

        }

    }
    
    if ($e3d_model -eq 'y') { & "$PSScriptRoot\02_E3D\03.01_Export_3D_model_from_E3D.ps1" }
   
    if ($fullrun -or $epc) {

        & "$PSScriptRoot\02_E3D\00.01_Export_Tags_from_E3D.ps1"

        $scripts = @("02.01_Export_parent_links_from_E3D.ps1",
                    "01.00_E3D_Tagged_Item_full_regex_Regex_Filtering.ps1",
                    "01.01_E3D_Non-Tagged_Items_Filtering.ps1",
                    "02_AIM_3D_model_links.ps1")
       
        foreach ($script in $scripts) {

            if ($packingvoke) {& "$PSScriptRoot\02_E3D\$script" -epc $epc}

                else {

                & "$PSScriptRoot\02_E3D\$script" -epc 11
                & "$PSScriptRoot\02_E3D\$script" -epc 12
                & "$PSScriptRoot\02_E3D\$script" -epc 13

                }
        
            }

        & "$PSScriptRoot\02_E3D\03.01_Export_3D_model_from_E3D.ps1"

    }

}

function E_I {

    param (
        [bool]$fullrun,     
        [bool]$packingvoke,     
        [string]$EI_tags,
        [string]$EI_props,
        [string]$epc
    ) 

    <# Tag export #>

    switch ($EI_tags) {

        "11" { & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 11  
            & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 12                             
            & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 13
        }        

    }
 
    <# Props export #>   

    switch ($EI_props) {

        "11" { & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 11  
            & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 12                             
            & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 13
        }        

    }
   
    if ($fullrun -or $epc) {

        $scripts = @("00_Extract_Tags_From_EI.ps1",
        "00_Extract_E&I_Properties_From_EI.ps1")

foreach ($script in $scripts) {

    if ($packingvoke) {& "$PSScriptRoot\03_EI\$script" -epc $epc}

    else {

        & "$PSScriptRoot\03_EI\$script" -epc 11
        & "$PSScriptRoot\03_EI\$script" -epc 12
        & "$PSScriptRoot\03_EI\$script" -epc 13

            }

        }

    }
}

function Engineering {

    param (
        [bool]$fullrun,      
        [bool]$packinvoke,      
        [string]$Eng_tags,
        [string]$Eng_props,
        [string]$epc
    ) 

    <# Tag export #>
  
    switch ($Eng_tags) {

        "11" { & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 11  
            & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 12                             
            & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 13
        }        

    }

    <# Properties export #>   

    switch ($Eng_props) {

        "11" { & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 11  
            & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 12                             
            & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 13
        }         

    }

    if ($fullrun -or $packingvoke) {

        $scripts = @("01.01_Export_Tags_From_Engineering.ps1",
                    "01.02_Export_Properties_from_Engineering.ps1")
 
        foreach ($script in $scripts) {

            if ($packingvoke) {& "$PSScriptRoot\04_Engineering\$script" -epc $epc}

            else {

            & "$PSScriptRoot\04_Engineering\$script" -epc 11
            & "$PSScriptRoot\04_Engineering\$script" -epc 12
            & "$PSScriptRoot\04_Engineering\$script" -epc 13

        }

        }

    }
}

function Diagrams {

    param (
        [bool]$fullrun,       
        [string]$Dia_tags
    )
    

    $scripts = @("01_Extract_Tags_from_Diagrams.ps1",
    "02_Extract_SCGROU_for_SVG_export.ps1")

    if ($Dia_tags -eq 'y') { foreach ($script in $scripts) { & "$PSScriptRoot\01_Diagrams\$script" } }
    
    if ($fullrun){ foreach ($script in $scripts) { & "$PSScriptRoot\01_Diagrams\$script" } }
  
}    
function Indexing {

    param (
        [bool]$fullrun,        
        [bool]$packingvoke,        
        [string]$meta_update,
        [string]$epc_envoke,
        [string]$aim_index,
        [string]$epc
    )
    
    #Data Fetching from MANASA

    $scripts = @("01.00_Delete_metadata_from_folder.ps1",
    "01.01_Extract_metadata_from_DMS.ps1",
    "01.02_Process_Metadata.ps1",
    "01.03_Extract_PDFs_from_MANASA.ps1")

    switch ($meta_update) {

        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13   

            }

        }     

    }

    <#Pure Indexing #>

    $scripts = @("01.04_Doc_indexing_multyTread.ps1",
    "01.05_Indexing_result_postProcessing.ps1")

    switch ($epc_envoke) {

        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13   

            }

        }        

    }

    <# AIM-A Section #>

    $scripts = @("02.01_Document_Register_for_AIM.ps1",
    "02.02_Publish_Doc_to_Tag.ps1",
    "02.03_PDF_copy_to_AIM.ps1")

    switch ($aim_index) {

       
        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13   

            }

        } 

    }

    if ($fullrun -or $packingvoke) {
       
        $scripts = @("01.00_Delete_metadata_from_folder.ps1",
            "01.01_Extract_metadata_from_DMS.ps1",
            "01.02_Process_Metadata.ps1",
            "01.03_Extract_PDFs_from_MANASA.ps1",
            "01.04_Doc_indexing_multyTread.ps1",
            "01.05_Indexing_result_postProcessing.ps1",
            "02.01_Document_Register_for_AIM.ps1",
            "02.02_Publish_Doc_to_Tag.ps1",
            "02.03_PDF_copy_to_AIM.ps1")

        foreach ($script in $scripts) {

            if ($packinvoke) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc $epc}

            else {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13
            
            }

        }
    }
}