param(
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [int]$epc
)


if ($epc -eq 13) {
    $filter = 'CPPR1-MDM5-AS*.svg'
}
if ($epc -eq 12) {
    $filter = 'RPBR1-LTE1-AS*.svg'
}
if ($epc -eq 11) {
    $filter = 'WHP*.svg'
}

$svg_list = Get-ChildItem -Path '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Diag\SVG\' -Filter $filter

$mapping_File = @{}
Import-Csv -Path "\\qamv3-sapp243\GDP\GDP_Config\Configuration\mapping.csv" | % {$mapping_File[$_."File name"] = $_."Document Identifier"}

foreach ($file in $svg_list) {
    $doc = $file.NAME
    
    if ($doc -notmatch '(CPPR1-MDM5-AS|RPBR1-LTE1-AS|WHPR1-MDM4-AS)[A-Z]{3}-[0-9]{2}-[0-9]{6}-[0-9]{4}_page[0-9].svg') {
        Write-Host "File skipped $doc"
        CONTINUE
    }
    
    if (-not($mapping_File[$doc])) {
        $doc + ',' + $doc -replace '_page[0-9].svg', ''  | Add-Content -Path "\\qamv3-sapp243\GDP\GDP_Config\Configuration\mapping.csv"
        Write-Host 'SVG added to the mapping file: $doc'
    }
}
Start-Sleep -Seconds 600
foreach ($file in $svg_list) {
    if ($doc -notmatch '(CPPR1-MDM5-AS|RPBR1-LTE1-AS|WHPR1-MDM4-AS)[A-Z]{3}-[0-9]{2}-[0-9]{6}-[0-9]{4}_page[0-9].svg') {
        CONTINUE
    }
    Copy-Item -Path $file.FullName -Destination '\\Qamv3-sapp243\gdp\GDP_StagingArea\NATIVE\Diagrams\' -Recurse -Force
    Write-Host "File copied for uploading into AIM-A $file"
}