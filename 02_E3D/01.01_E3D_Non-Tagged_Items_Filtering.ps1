<#

    Purpose: Filter out "Non-tagged" objects based on RegEx from Class Library. Preparing
    Loading file for the AIM-A.
#>


param(
    [Parameter(Mandatory=$true)]
    [ValidateSet(11, 12, 13)]
    [string] $epc

)

# Get-Job | Remove-Job


$files = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged" -Filter "EPCIC$epc*-parents.csv"
# $files = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged" -Filter "EPCIC$epc*-parents.csv"

# === Load XML and extract regex patterns ===
$ism_file = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\RegisterGateway\source data\class library" -Filter clib.xml | Sort-Object | Select-Object -Last 1

[xml]$xml = Get-Content $ism_file.FullName
$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsMgr.AddNamespace("a", "http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01")
 
$nodes = $xml.SelectNodes("/a:ClassLibrary/a:ReferenceData/a:NamingAndNumbering/a:Templates/a:Template[contains(@name, 'Non_tagged')]", $nsMgr)


# $regexPatterns = @()
$joined = @()
foreach ($node in $nodes) {
    $pattern = $node.GetAttribute("regEx")
    $class = $node.GetAttribute("id").replace('Non_tagged_items_','')
    if ($class) {
        $part = $class + '@' + $pattern
        $joined += $part
    }
}

foreach ($file in $files) {

    $clean = $file.Name -replace ('_E3D-parents.csv', '')

    $data = Import-Csv -Path $file.FullName -Delimiter ';'

$totalRows = $data.Count
$jobcount = 5
$chunkSize = [Math]::Ceiling($totalRows / $jobcount)
$chunks = New-Object System.Collections.Generic.List[object]

# === Chunk the name values ===

for ($i = 0; $i -lt $totalRows; $i += $chunkSize) {
    $end = [math]::Min($i + $chunkSize - 1, $totalRows - 1)
    $chunks.Add($data[$i..$end])
}

# === Define job script block ===
$scriptBlock = {
    # param($rows, $regexList, $classList)
    param($rows, $regexList)

    $results = @()
    foreach ($row in $rows) {
		$item = $row.NAME
        
    foreach ($part in $regexList) {
            $rg = $part -replace('.+@','')
            $cl = $part -replace('@.+','')
            if ($item -match $rg) {
                $results += [pscustomobject]@{
                    Model = $item.Substring(1,4) + '_3D_MODEL'
                    Tag_Number = $item -replace('/','')
                    Tag_Class = $cl
                    Tag_Description = 'Non Tagged Item'
                    Ref3D = $item
                    Status = $null
                    Action = $null
                    Platform = $item.Substring(1,4)
                }
                break
            }
        }
    }
    return $results
}

$jobs = @()
foreach ($chunk in $chunks) {
    $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList @($chunk, $joined)
    $jobs += $job
}

if ($jobs.Count -eq 0) {
    Write-Error "No jobs were created. Aborting."
    return
}

# === Wait and gather results ===
Wait-Job -Job $jobs

$allMatches = @()
foreach ($job in $jobs) {
    $jobResults = Receive-Job -Job $job
    if ($jobResults) {
        $allMatches += $jobResults
    }
    Remove-Job -Job $job
}

$root_path = "\\Qamv3-sapp243\gdp\GDP_StagingArea\MP\E3D_TAGS\"
#$root_path = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\AIMA\non-tagged\"
$outputPath = $root_path + $clean + "-E3D_Non_Tagged_AIM.csv"


# === Export results to CSV ===
$allMatches | Select-Object Model, Tag_Number, Tag_Class, Tag_Description, Ref3D, Status, Action, Platform  |
    Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Delimiter ','
}