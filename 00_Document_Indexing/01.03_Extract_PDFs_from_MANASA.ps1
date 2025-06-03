param (
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [string] $epc
)

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "INDEXING"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $scriptname. Please Wait"
Write-Log -Level INFO -Message "====================================="

Remove-Item -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\logs\*" -Recurse
# $date = Get-Date -Format "yyyy-MM-dd"
# $log =  "W:\Appli\DigitalAsset\MP\RUYA_data\Logs\PS\" + $date + "_01_Update_Dcument_rendition.log"
# Start-Transcript -Path $log -Append

$path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC"+$epc+"_Output.xml"

$cmis_config_path = "D:\CMISGateway\NOC-RUYA\EPC"+ $epc +"_Indexing.xml"
Start-Process -Filepath "C:\Program Files\AVEVA\AVEVA NET Gateways\Gateway For CMIS\AVEVA.NET.Gateways.CMIS.App.exe" `
-ArgumentList "-cp $cmis_config_path -un cgateway -pw AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAA51vIVy3Nk0W15ZGhohO1GQQAAAACAAAAAAADZgAAwAAAABAAAABaHe9YH4W0rxGXKiFJ5dNmAAAAAASAAACgAAAAEAAAAHogP97OLajdPkwy9ViTKtgoAAAAIL4/ts71sa/zp7+JaAiOOPFZ0jsZzS2NwcgEhFZWScmhS4cMPgybDxQAAAC3Z32+R/8+eNJzMrMGVFDtYYKPFw== -ol $path" -NoNewWindow -Wait

$source_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\indexing\EPC"+$epc+"_Source"

# Load .NET ZIP support once
Add-Type -AssemblyName System.IO.Compression.FileSystem

do {
    $zips = Get-ChildItem -Path $source_path -Filter *.zip -Recurse
    foreach ($zip in $zips) {
        # Write-Host "Processing $($zip.FullName)"
        $archive = [IO.Compression.ZipFile]::OpenRead($zip.FullName)
        try {
            $archive.Entries |
              Where-Object {
                  $ext = [IO.Path]::GetExtension($_.FullName).ToLowerInvariant()
                  $ext -in '.dwg','.xlsx','.xls','.zip'
              } |
              ForEach-Object {
                  $entry = $_
                  $dest = Join-Path $zip.DirectoryName $entry.Name
                #   Write-Host " Extracting $($entry.FullName) → $dest"
                  [System.IO.Compression.ZipFileExtensions]::ExtractToFile(
                      $entry,
                      $dest,
                      $true
                  )
              }
        }

        catch {

            # Write-Log -Level INFO -Message "Encountered problem with the file $($zip.FullName)"
            Write-Host -Level INFO -Message "Encountered problem with the file $($zip.FullName)"
        }

        finally {
            $archive.Dispose()
        }

        Remove-Item -LiteralPath $zip.FullName -Force
    }
} while ($zips.Count -gt 0)

########## NEW ZIP LOGIC ##############################

$inArray_source = [System.IO.Directory]::GetFiles("$source_path" , "*.pdf", [System.IO.SearchOption]::AllDirectories).Length


#Checking for potential error

$files = Get-ChildItem -Path "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\logs\*"

foreach ($file in $files) {
    # Get-Content -Path $filef
    $content = Get-Content -Path $file
    if ($content -match '\[Error\]') {
        Write-Log -Level ERROR -Message "Error for the file $file"
        Write-Log -Level ERROR -Message "$content"
        Write-Error Get-Content -Path $file
    }
}


Write-Log -Level INFO -Message "Total count of PDF files in the Source Folder: $inArray_source"
$finished = $true
Write-Log -Level INFO -Message "Extraction Completed" -finished $finished
