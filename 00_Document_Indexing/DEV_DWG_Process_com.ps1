param(
    # [Parameter(Mandatory)][string] $SourceDir,
    # [Parameter(Mandatory)][string] $OutCsv,
    [string] $RegexCsv = 'W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv',
    [string[]] $InstList
)

# $SourceDir = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\CPPR1-MDM5-ASBJA-10-R54062-0001\03"
$SourceDir = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\Indexing\smallbatch"
$OutCsv = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\Indexing\out.csv"

#ensure we are in STA (COM req)

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw "Invoke DwgIndexingCom must run in a SINGLE THREADED apartment. " +
          "Launch the worker run space with ApartmentState = 'STA'."
}


if (-not $InstList) {
    $InstList = @(
"AAH","AAHH","AI","AIS","AIT","APB","AR","ARC","ASP","AT","AV","BDV","BLD","BPR","BX","BY","CAM","CC","CHV","CI","CMO","CMP","CPF","CPJ","CPR","CS","CTP","CVA","CY","DI","DRS","DT","EJX","EPB","EPR","ESDV","EWS","EX","EY","FA","FAH","FAHH","FAL","FALL","FC","FCS","FCV","FE","FG","FHA","FI","FIT","FIV","FMX","FO","FPS","FQ","FQI","FQV","FQVY","FS","FSH","FSHH","FSL","FSLL","FT","FVI","FVS","FX","FY","GD","GDAH","GDAHH","GDR","GDS","GDT","GLV","GVA","GVAA","HC","HCS","HCV","HD","HDAH","HDAHH","HDC","HDR","HDS","HDT","HF","HG","HGAH","HGAHH","HGS","HIT","HR","HRAH","HS","HSS","HT","HVA","HVS","IAM","ICD","ID","ILK","IMS","IPC","IR","IRAH","JBC","JBE","JBF","JBJ","JBS","LAH","LAHH","LAL","LALL","LC","LCV","LG","LI","LIT","LOS","LRS","LS","LSC","LSD","LSH","LSHH","LSHL","LSL","LSLL","LSS","LT","LVI","LY","MAC","MACA","MCT","MCV","MI","MOV","MRD","MT","MWS","OCP","OWS","PA","PAH","PAHH","PAL","PALL","PB","PC","PCD","PCV","PDAH","PDAHH","PDAL","PDALL","PDC","PDCV","PDI","PDIT","PDRC","PDS","PDSH","PDSHH","PDSL","PDSLL","PDT","PDY","PE","PI","PIT","PRI","PRS","PRV","PS","PSE","PSH","PSHH","PSL","PSLL","PSV","PT","PV","PVI","PX","PY","R","RCU","RD","RO","RTD","RTU","S","SAH","SAHH","SAL","SALL","SCP","SD","SDAH","SDV","SE","SI","SL","SOV","SPR","SS","SSH","SSL","SSSV","ST","SVC","SVP","SWS","SX","SY","TAH","TAHH","TAL","TALL","TC","TCV","TDAH","TDAL","TDIC","TDY","TE","TES","TI","TIT","TMX","TS","TSH","TSHH","TSHL","TSL","TSLL","TSV","TT","TVI","TW","TY","UA","UV","VAH","VAHH","VDU","VGDAH","VGDAHH","VHDAH","VHDAHH","VHGAH","VHGAHH","VHRAH","VIRAH","VMACA","VPSV","VSDAH","VT","WAA","WCV","WMA","WMH","WML","WMR","WMV","WT","X","XA","XAH","XAHH","XC","XCT","XCV","XEP","XI","XL","XPI","XPS","XS","XT","XY","Y","YSL","ZAH","ZAHH","ZE","ZI","ZIC","ZIO","ZL","ZLC","ZLO","ZS","ZSC","ZSO","ZT","2WV","3WV","AAL","ABE","ABI","ABIT","ABT","AC","ACUSH","ACUSL","ADTN","AEN","AGE","AGI","AGIT","AGT","AH","AHA","AHE","AHI","AHIT","AHS","AHT","AIC","AIN","AITN","AME","AMI","AMIT","AMT","AO","AOJ","AP","APCSH","APCSL","ART","ARV","ASCSH","ASCSL","ASH","ASHL","ASL","ASPSH","ASPSL","ATM","ATN","AVY","AWT","AX","AY","AZ","AZE","AZI","AZIT","AZL","AZP","AZR","AZS","AZSC","AZSO","AZT","AZTN","BAL","BDIM","BDIOM","BDOM","BFCL","BFV","BG","BI","BIAD","BIAL","BIALS","BIAS","BIC","BIT","BL","BP","BPV","BR","BRC","BRT","BS","BSG","BSH","BSHL","BSL","BSP","BT","BTC","BTF","BTH","BTHA","BTHL","BTHLI","BTHLR","BTK","BTKL","BTM","BW","BZE","BZL","BZP","BZR","BZS","BZW","CAH","CAL","CE","CGE","CGT","CGTN","CIT","CSC","CSH","CSL","CSO","DE","DHS","DO","DPSH","DPSL","DR","DTT","DX","DY","DZY","EE","EG","EI","EIC","EIT","ER","ERC","ERT","ESD","ESH","ESHL","ESL","ET","EZE","EZI","EZL","EZP","EZR","EZW","FBE","FBI","FBIT","FBS","FBT","FDTN","FEN","FFC","FFE","FFI","FFIC","FFR","FFRC","FFSH","FFSL","FGE","FGI","FGIT","FGN","FGS","FGT","FHE","FHG","FHI","FHIT","FHS","FHT","FICV","FIN","FITN","FITNSH","FITNSL","FJB","FL","FM","FME","FMI","FMIT","FMS","FMT","FQE","FQG","FQIC","FQIT","FQR","FQRC","FQSH","FQSL","FQT","FQX","FQY","FR","FRC","FRT","FSHL","FSV","FTN","FU","FV","FVY","FZ","FZE","FZI","FZIT","FZL","FZP","FZR","FZS","FZSC","FZSO","FZT","FZTN","FZU","FZV","FZW","FZY","FZZS","FZZSC","FZZSO","GAE","GAT","GATN","GAV","GBE","GBT","GBTN","GCE","GCFT","GCFTN","GCT","GCTN","GDE","GDTN","GEE","GET","GETN","GFE","GFT","GFTN","GGAE","GGAT","GGATN","GGBE","GGBT","GGBTN","GGCE","GGCT","GGCTN","GGDE","GGDT","GGDTN","GGE","GGEE","GGET","GGETN","GGFE","GGFT","GGFTN","GGGE","GGGT","GGGTN","GGHE","GGHT","GGHTN","GGIE","GGIT","GGITN","GGJE","GGJT","GGJTN","GGKE","GGKT","GGKTN","GGLE","GGLT","GGLTN","GGME","GGMT","GGMTN","GGNE","GGNT","GGNTN","GGOT","GGOTN","GGPT","GGPTN","GGRT","GGRTN","GGST","GGSTN","GGT","GGTN","GHCT","GHCTN","GHE","GHLT","GHLTN","GHT","GHTN","GIE","GIT","GITN","GJE","GJT","GJTN","GKE","GKT","GKTN","GLE","GLT","GLTN","GMCE","GMCT","GMCTN","GME","GMT","GMTN","GNE","GNOE","GNOT","GNOTN","GNT","GNTN","GOE","GOT","GOTN","GPE","GPOT","GPOTN","GPT","GPTN","GRE","GRT","GRTN","GSE","GST","GSTN","GTT","GTTN","GUE","GUT","GUTN","GVE","GVT","GVTN","GWE","GWT","GWTN","GXE","GXT","GXTN","GYE","GYT","GYTN","GZE","GZT","GZTN","HBB","HBS","HBV","HBY","HHS","HIC","HL","HOA","HSA","HSC","HSO","HSZ","HV","HVY","HX","HY","HYC","HYO","HZIC","HZIO","HZR","HZS","HZSC","HZSO","HZV","HZY","HZZS","HZZSC","HZZSO","IA","II","IIC","IIT","IL","IP","IRC","IRT","ISH","ISHL","ISL","IT","IV","IX","IY","IZE","IZI","IZL","JE","JG","JI","JIC","JIT","JO","JOI","JOR","JQX","JR","JRC","JRT","JSH","JSHL","JSL","JT","JX","JY","JZE","JZI","JZL","JZP","JZR","K","KC","KCV","KE","KG","KI","KIC","KIT","KL","KME","KOG","KOL","KOX","KPE","KQI","KR","KRC","KRT","KSH","KSHL","KSL","KT","KV","KX","KY","KZSC","KZSD","KZSO","KZSS","KZT","KZV","KZY","KZZSC","KZZSO","LA","LACT","LAD","LAHI","LBE","LBI","LBIT","LBP","LBS","LBT","LCI","LCVI","LD","LDI","LDIN","LDIT","LDT","LDTN","LDZT","LDZU","LEI","LEN","LGE","LGI","LGIT","LGS","LGT","LHE","LHG","LHI","LHIT","LHS","LHT","LICI","LIN","LINSH","LINSL","LITN","LITNSH","LITNSL","LL","LLC","LLG","LME","LMI","LMIT","LMS","LMT","LO","LR","LRC","LSHI","LSLI","LTI","LTN","LUX","LV","LVS","LVY","LW","LWG","LX","LZE","LZI","LZIN","LZIT","LZL","LZP","LZS","LZSC","LZSH","LZSL","LZSN","LZSO","LZT","LZTN","LZTNSH","LZTNSL","LZV","LZY","MA","MBOV","MBZSC","MBZSO","MCU","MHHS","MHS","MHT","MLPSH","MLPSL","MZHS","MZLC","MZLO","MZSC","MZSO","MZZS","MZZSC","MZZSO","MZZT","NO","NOC","OM","OV","OWD","OX","PAD","PBI","PBIT","PBS","PBT","PCVN","PDBI","PDBIT","PDBS","PDBT","PDE","PDG","PDGI","PDGIT","PDGN","PDGS","PDGT","PDHCV","PDHG","PDHI","PDHIT","PDHS","PDHSH","PDHSL","PDHT","PDIN","PDITN","PDMI","PDMIT","PDMS","PDMT","PDR","PDRT","PDTN","PDV","PDVY","PDX","PDZI","PDZIT","PDZS","PDZSC","PDZSH","PDZSO","PDZT","PDZTN","PEN","PFL","PFR","PFX","PG","PGI","PGIT","PGN","PGS","PHG","PHI","PHIT","PHS","PHT","PHV","PIN","PITN","PJR","PK","PKL","PKR","PKX","PLC","PLPSH","PLPSL","PMI","PMIT","PMS","PMT","PN","POL","PP","PRC","PRT","PSHL","PSVNSH","PSVNSL","PTC","PTN","PVN","PVY","PZE","PZI","PZIN","PZIT","PZL","PZP","PZS","PZSC","PZSH","PZSL","PZSO","PZT","PZTN","PZTNSH","PZTNSL","PZV","PZY","PZZS","PZZSC","PZZSO","QE","QI","QIC","QIT","QOL","QQ","QQI","QQR","QQX","QR","QRC","QRT","QSH","QSHL","QSL","QT","QY","RAV","RE","REG","RI","RIC","RIT","ROR","ROX","RP","RQ","RQI","RQL","RRC","RRT","RSH","RSHL","RSL","RT","RW","RY","RZE","RZL","RZP","SB","SC","SCN","SCNSH","SCNSL","SCV","SG","SHT","SIC","SIT","SJ","SLO","SME","SMT","SP","SR","SRC","SRT","SRV","SSHL","SSV","SV","SVY","SZE","SZIT","SZT","SZY","TBE","TBI","TBIT","TBS","TBT","TD","TDA","TDC","TDCV","TDE","TDG","TDI","TDIT","TDL","TDR","TDRC","TDRT","TDS","TDSH","TDSL","TDT","TDTN","TDV","TDVY","TDX","TEN","TFI","TFR","TFX","TG","TGE","TGI","TGIT","TGS","TGT","THCV","THE","THG","THI","THIT","THS","THT","THV","TIN","TINSH","TINSL","TIS","TISHL","TITN","TJ","TJE","TJX","TK","TKR","TKX","TL","TME","TMI","TMIT","TMS","TMT","TOR","TP","TR","TRC","TRS","TRT","TSE","TTN","TV","TVY","TZE","TZI","TZIT","TZL","TZP","TZR","TZS","TZSC","TZSH","TZSL","TZSO","TZT","TZTN","TZV","TZW","TZY","UE","UI","UJ","UJR","UL","ULO","UR","USD","UTN","UVY","UX","UY","UZ","UZE","UZL","UZP","UZR","UZS","UZSC","UZSO","UZW","UZY","VBE","VBIT","VBS","VBT","VE","VFD","VG","VGE","VGIT","VGS","VGT","VHE","VHIT","VHS","VHT","VI","VIT","VL","VME","VMIT","VMS","VMT","VP","VR","VRT","VS","VTA","VX","VXE","VXG","VXI","VXL","VXME","VXR","VXT","VXX","VY","VYE","VYME","VYP","VYT","VZ","VZE","VZG","VZIT","VZL","VZP","VZR","VZS","VZT","VZX","WA","WAI","WAR","WAX","WC","WDC","WDCV","WDI","WDIC","WDIT","WDL","WDR","WDRC","WDRT","WDSH","WDSL","WDT","WDX","WE","WEC","WEN","WFI","WFR","WFX","WG","WHS","WI","WIC","WIN","WIT","WKI","WKR","WKX","WL","WMS","WQI","WQL","WQR","WQX","WR","WRT","WS","WSH","WSHL","WSL","WTN","WUL","WUP","WUR","WX","WXE","WXI","WXL","WXR","WXX","WY","WYE","WYL","WZ","WZE","WZIT","WZL","WZR","WZS","WZT","WZX","WZXL","WZXP","WZXR","WZYL","WZYP","WZYR","XBY","XG","XHY","XHZSC","XIO","XLC","XLO","XP","XQ","XR","XV","XVN","XVY","XW","XX","XYC","XYD","XYH","XYL","XYO","XYS","XZC","XZIC","XZIO","XZLC","XZLO","XZO","XZS","XZSC","XZSD","XZSH","XZSL","XZSO","XZSP","XZSS","XZV","XZY","XZYC","XZYO","XZZLC","XZZLO","XZZS","XZZSC","XZZSD","XZZSH","XZZSL","XZZSO","XZZSS","YC","YE","YI","YIC","YL","YR","YSH","YT","YX","YY","YZ","YZE","YZL","YZR","YZX","ZC","ZCV","ZDC","ZDCV","ZDE","ZDG","ZDI","ZDIC","ZDIT","ZDL","ZDR","ZDRC","ZDRT","ZDSH","ZDSL","ZDT","ZDX","ZDXE","ZDXG","ZDXI","ZDXL","ZDXR","ZDXX","ZDY","ZDYE","ZDYG","ZDYI","ZDYL","ZDYR","ZDYX","ZDZE","ZDZI","ZDZL","ZDZR","ZDZX","ZG","ZHS","ZHSC","ZHSO","ZHT","ZIT","ZME","ZMT","ZO","ZP","ZR","ZRC","ZRT","ZSD","ZSE","ZSH","ZSHL","ZSL","ZSP","ZST","ZTN","ZUI","ZUL","ZUP","ZUR","ZW","ZX","ZXE","ZXG","ZXI","ZXL","ZXR","ZY","ZYE","ZYG","ZYL","ZYR","ZYX","ZZ","ZZI","ZZX","ZZXE","ZZXI","ZZXL","ZZXP","ZZXR","ZZYE","ZZYI","ZZYL","ZZYP","ZZYR")
}


if (-not (Test-Path $RegexCsv)) {
    throw "Light_regex.csv not found → $RegexCsv"
}
$LightRows = Import-Csv -Delimiter ';' -Path $RegexCsv
$LightCompiled = foreach ($row in $LightRows) {
    [regex]::new(($row.Regexp -replace '\$$','(,|;)?$'),'Compiled,IgnoreCase')
}

$InstPat = ($InstList -join '|')
$RSeq = [regex]::new("^($InstPat)-[0-9]{6}$",'Compiled,IgnoreCase')

New-Item -Path (Split-Path $OutCsv) -ItemType Directory -Force | Out-Null

#Dedupe
$seen = New-Object System.Collections.Generic.HashSet[string]
$today = Get-Date -Format 'MM/dd/yyyy'

# Start BricsCAD (hidden)
$brics = New-Object -ComObject BricscadApp.AcadApplication
$brics.Visible = $false

try {

    $dwgFiles = Get-ChildItem -Path $SourceDir -Recurse -Include *.dwg -File
    Write-Host "DWG COM: processing $($dwgFiles.Count) files"

    foreach ($file in $dwgFiles) {
        try { $doc = $brics.Documents.Open($file.FullName,$false) }
        catch { Write-Warning "Open failed : $($_.Exception.Message)"; continue }

        $rows = New-Object System.Collections.Generic.List[psobject]

        $ms = $doc.ModelSpace

        #Collect all circles once per dtawing (centre + radius)

        $circles = @()

        for ($i = 0; $i -lt $ms.Count; $i++) {
            $ent = $ms.item($i)
            if ($ent.ObjectName -eq 'AcDbCircle') {
                $center = [object[]]$ent.Center
                $radius = [double]$ent.Radius
                $circles += [pscustomobject]@{
                    Cx = [double]$center[0]
                    Cy = [double]$center[0]
                    R2 = $radius * $radius
                }
            }
        }

        #Iterate text/mtext/attributes

        for ($i = 0; $i -lt $ms.Count; $i++) {
            $ent = $ms.item($i)


            switch ($ent.ObjectName) {
                'AcDbText' {$txt = $ent.TextString; $pt = [object[]]$ent.InsertionPoint}
                'AcDbMText' {$txt = $ent.TextString; $pt = [object[]]$ent.InsertionPoint}
                'AcDbAttributeReference' {$txt = $ent.TextString; $pt = [object[]]$ent.InsertionPoint}
                default {continue}
            }

            if ([string]::IsNullOrWhiteSpace($txt)) {continue}

            $inCircle = $false
            foreach ($c in $circles) {
            #distance between text insertion and circle centre
                $dx = [double]$pt[0] - $c.Cx
                $dy = [double]$pt[1] - $c.Cy
                if (($dx*$dx + $dy*$dy) -le $c.R2) {$inCircle =$true; break}
            }

            foreach ($token in ($txt -split '\s*[;,]\s*' | Where-Object{$_ -and $_.Length -le 60})){

                $key = "$token|$($file.BaseName)"
                if ($seen.Contains($key)) { continue }

                # Light regex
                $hit = $false
                for ($x=0;$x -lt $LightCompiled.Count;$x++) {
                    if ($LightCompiled[$x].IsMatch($token)) {
                        $rows.Add([pscustomobject]@{
                            Tag_number = $token
                            Document_number = $file.BaseName
                            ST = $LightRows[$i].Naming_template_ID
                            DATE = $today
                            file_full_path = $file.FullName
                            SourceType = 'DWG'
                        })
                        $seen.Add($key) | Out-Null
                        $hit = $true ; break
                    }
                }
                if ($hit) { continue }

                # Instrument Tag (prefix-seq)
                if ($RSeq.IsMatch($token)) {
                    $rows.Add([pscustomobject]@{
                        Tag_number = $token
                        Document_number = $file.BaseName
                        ST = 'Hand valve custom search'
                        DATE = $today
                        file_full_path = $file.FullName
                        SourceType = 'DWG'
                    })
                    $seen.Add($key) | Out-Null
                    continue
                }

                # Circle-enclosed tags
                if ($inCircle){
                    $rows.Add([pscustomobject]@{
                        Tag_number = $token
                        Document_number = $file.BaseName
                        ST = 'Circle Tag'
                        DATE = $today
                        file_full_path = $file.FullName
                        SourceType = 'DWG'
                    })
                    
                    $seen.Add($key) | Out-Null
                }
            }
        }
      $doc.Close($false)

    } 
} finally {
    
    $brics.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($brics)
}

$rows | Export-Csv -path $OutCsv -NoTypeInformation -Encoding UTF8 -Force

Write-Host "DWG COM: finished at $OutCsv"
