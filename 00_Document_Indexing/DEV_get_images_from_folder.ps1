# Define paths
$pdfPath = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\CPPR1-MDM5-ASBJA-10-PV4023-0001\00\CPPR1-MDM5-ASBJA-10-PV4023-0001.pdf"
$outputFolder = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\CPPR1-MDM5-ASBJA-10-PV4023-0001\00\"



$local_path = $PSScriptRoot
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'UglyToad.PdfPig.dll')
Import-Module $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'UglyToad.PdfPig.DocumentLayoutAnalysis.dll')
Import-Module $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'BouncyCastle.Crypto.dll')
Import-Module $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'itextsharp.dll')
Import-Module $dll_path
Add-Type -Path $dll_path
$dll_path = [System.IO.Path]::Combine($local_path,'lib', 'itextsharp.pdfa.dll')
Import-Module $dll_path



$pdf =  [UglyToad.PdfPig.PdfDocument]::Open($pdfPath)


$page = $pdf.GetPage(10)
$letters = $page.Letters #| Where-Object {
    #$_.Value -match '^[0-9A-Za-z\-]$'
#} #| Sort-Object -Property @{Expression = { $_.Location.X }}

$words = @()
$currentWord = ""
$lastLetter = $null

foreach ($letter in $letters) {
    # if ($lastLetter -eq $null) {
    #     $currentWord = $letter.Value
    # } else {
    #     $gap = $letter.Location.X - ($lastLetter.Location.X + $lastLetter.Width)
    #     $threshold = $lastLetter.Width * 0.9

    #     if ($gap -le $threshold) {
    #         $currentWord += $letter.Value
    #     } else {
    #         $words += $currentWord
    #         $currentWord = $letter.Value
    #     }
    # }
    # $lastLetter = $letter
    $currentWord += $letter.Value
}

if ($currentWord -ne "") {
    $words += $currentWord
}

Write-Host "Page $pageNum Words:"
$words | ForEach-Object { Write-Host $_ }


$pdf.Dispose()

