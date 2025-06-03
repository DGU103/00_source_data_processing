param (
    [Parameter(Mandatory=$true)]
    [ValidateSet('11','12','13')]
    [String]$epc
)


if ($epc -eq '13') {
    $package_name = "EPC13"
}
elseif ($epc -eq '12') {
    $package_name = "EPC12"
}
elseif ($epc -eq '11') {
    $package_name = "EPC11"
}

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "AIM"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "Running $scriptname for EPCIC $epc"
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"

if($epc -in @('11','12','13')){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\"}
elseif($epc -eq '06'){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\CPP03\Source\Indexing\"}
elseif($epc -eq '05'){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\"}

$doc_source_folder = $root_path+$package_name + "_Source\"

Write-Log -Level INFO -Message "Collecting PDF files from the Source folder..."

$pdf_files = Get-ChildItem -Path $doc_source_folder -Filter *.pdf -Recurse 

$count = $pdf_files.Count

Write-Host ("")
# Write-Host ("Collected PDF files count is $count") -ForegroundColor Cyan
Write-Log -Level INFO -Message "Collected PDF files count is $count"


$objects = @()
$objects += "File name,Document Identifier"
<# This section used to copy only new documents from source #>
Write-Log -Level INFO -Message ("Reading CSV mapping file from \\qamv3-sapp243\GDP\GDP_Config\Configuration\mapping.csv")

$temp = Import-Csv -Path "\\qamv3-sapp243\GDP\GDP_Config\Configuration\mapping.csv" 
$mapping_file =@{}

Write-Log -Level INFO -Message "Processing CSV.."
foreach ($t in $temp) {
   if (-not ($mapping_file[$t."File name"])) {
      $mapping_file.Add($t."File name", $t."Document Identifier")
      # Write-Log -Level DEBUG -Message "Processing and Adding info into CSV:    $t"
   }
}
Write-Log -Level INFO -Message "Updating ARRAY.."
foreach ($pdf_file in $pdf_files) {
   if (-not ($mapping_file[$pdf_file.Name])) {
      $mapping_file.Add($pdf_file.Name, $pdf_file.Name.Replace(".pdf",""))
      # Write-Log -Level DEBUG -Message "Updating the ARRAY for $pdf_file"
   }
}
foreach ($key in $mapping_file.Keys) {
   $objects += $key + ',' + $mapping_file[$key]
}

$objects | Out-File -FilePath "\\qamv3-sapp243\GDP\GDP_Config\Configuration\mapping.csv"
$destination_path = "\\Qamv3-sapp243\gdp\GDP_StagingArea\NATIVE\EDMS"

Write-Log -Level INFO -Message "Beginning copy of new documents into GDP Staging Area: $destination_path"

foreach ($pdf_file in $pdf_files) {

    $relativePath = $pdf_file.FullName.Substring($doc_source_folder.Length).TrimStart('\')

    # Construct the corresponding path in the destination
    $destinationFilePath = Join-Path $destination_path $relativePath

    # 1) Check if the file with the same revision already exists
    if (Test-Path $destinationFilePath) {
        Write-Log -Level DEBUG -Message "SKIPPING: $($pdf_file.FullName); same revision already in $destinationFilePath"
        continue
    }

    # 2) If it does not exist, ensure the destination subfolder is created
    $destDir = Split-Path -Path $destinationFilePath -Parent
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    # 3) Copy the file
    Write-Log -Level INFO -Message "Copying Over the files into Staging Area. Please Wait."
    Copy-Item -Path $pdf_file.FullName -Destination $destinationFilePath
    Write-Log -Level DEBUG -Message "COPIED: $($pdf_file.FullName) -> $destinationFilePath"
}

$finished = $true
Write-Log -Level INFO -Message "Copy operation completed." -finished $finished