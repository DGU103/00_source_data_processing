# Shortcut .NET types
using namespace Autodesk.AutoCAD.DatabaseServices
using namespace Autodesk.AutoCAD.ApplicationServices
using namespace Autodesk.AutoCAD.Runtime

param(
    # [Parameter(Mandatory)][string] $SourceDir,
    # [Parameter(Mandatory)][string] $OutCsv,
    [string] $RegexCsv = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Light_regex.csv'
)

$sourcedir = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\CPPR1-MDM5-ASBJA-10-R54062-0001\03"
$outcsv = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\dwgtest.csv"

# ─── Load BricsCAD .NET assemblies once per PowerShell session ────────────
$brxRoot = 'C:\Program Files\Bricsys\BricsCAD V24 en_US'

# foreach ($dll in @('acdbmgd.dll','acmgd.dll','BrxMgd.dll')) {
foreach ($dll in @('BrxMgd.dll')) {
    $path = Join-Path $brxRoot $dll
    if (-not (Get-Module -ListAvailable | Where-Object {$_.Path -eq $path})) {
        Add-Type -Path $path -ErrorAction Stop
    }
}

# ─── Compile Light Regex patterns ──────────────
$Light_Regex = Import-Csv -Delimiter ';' -Path $RegexCsv
$LightRegexCompiled = foreach ($row in $Light_Regex) {
    $p = $row.Regexp -replace '\$$','(,|;)?$'
    [regex]::new($p, 'Compiled,IgnoreCase')
}

# ─── Collector arrays (shared for the whole folder) ─────────
$tags = @()
$date = Get-Date -Format 'MM/dd/yyyy'

# ─── Iterate *.dwg (single thread; parallelism comes from launcher) ───────
$dwgFiles = Get-ChildItem -Path $SourceDir -Recurse -Include *.dwg -File
Write-Host "DWG pipeline: found $($dwgFiles.Count) files in $SourceDir"

foreach ($file in $dwgFiles) {

    try {

        $doc = [Application]::DocumentManager.Open($file.FullName, $false)
        $db = $doc.Database

        $tr = $db.TransactionManager.StartTransaction()
        try {
            $btr = [BlockTableRecord] $tr.GetObject(
                       $db.CurrentSpaceId,
                       [OpenMode]::ForRead
                   )

            foreach ($id in $btr) {
                $ent = $tr.GetObject($id, [OpenMode]::ForRead)

                # ---- Plain text ----------------------------------------------------
                if ($ent -is [DBText] -or $ent -is [MText]) {
                    $txt = $ent.TextString
                    for ($i = 0; $i -lt $LightRegexCompiled.Count; $i++) {
                        if ($LightRegexCompiled[$i].IsMatch($txt)) {
                            $rx = $Light_Regex[$i]
                            $t = [pscustomobject]@{
                                Tag_number = $txt
                                Document_number = $file.BaseName
                                doctitle = $null
                                doctype = 'DWG'
                                issuance_code = $null
                                ST = $rx.Naming_template_ID
                                DATE = $date
                                doc_date = $null
                                issue_reason = $null
                                file_full_path = $file.FullName
                            }
                            $tags += $t
                            break
                        }
                    }
                }

                # ---- Block attribute values ---------------------------------------
                elseif ($ent -is [BlockReference]) {
                    foreach ($attId in $ent.AttributeCollection) {
                        $att = $tr.GetObject($attId, [OpenMode]::ForRead)
                        $txt = $att.TextString
                        for ($i = 0; $i -lt $LightRegexCompiled.Count; $i++) {
                            if ($LightRegexCompiled[$i].IsMatch($txt)) {
                                $rx = $Light_Regex[$i]
                                $t = [pscustomobject]@{
                                    Tag_number = $txt
                                    Document_number = $file.BaseName
                                    doctitle = $null
                                    doctype = 'DWG'
                                    issuance_code = $null
                                    ST = $rx.Naming_template_ID
                                    DATE = $date
                                    doc_date = $null
                                    issue_reason = $null
                                    file_full_path = $file.FullName
                                }
                                $tags += $t
                                break
                            }
                        }
                    }
                }
            } # foreach entity
            $tr.Commit()
        }
        finally {
            $tr.Dispose()
            [Application]::DocumentManager.MdiActiveDocument.CloseAndDiscard()
        }
    }
    catch {
        Write-Warning "DWG parse failed: $($file.Name) - $($_.Exception.Message)"
    }
}

# ─── Export once for the whole folder -------------------------------------
if ($tags) {
    $tags | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "DWG pipeline wrote $($tags.Count) rows → $OutCsv"
}
else {
    Write-Host "DWG pipeline - no matches."
}
