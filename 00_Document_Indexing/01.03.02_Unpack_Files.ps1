<#
.SYNOPSIS
    Recursively unpack *.zip archives (including zips in zips) for the following types: DWG, XLS and XLSX
    
    Move any archive that cannot be processed to a “Corrupted_files” folder.

#>

param (
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [string] $epc
)

$SourcePath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\indexing\EPC"+$epc+"_Source"
$CorruptedPath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Corrupted_files"

#Make sure .NET's ZipFile class is available only once.
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

# Ensure the quarantine folder exists.
if (-not (Test-Path $CorruptedPath)) {
    New-Item -ItemType Directory -Path $CorruptedPath | Out-Null
}

function Expand-ArchiveRecursively {
    param(
        [IO.FileInfo]$ZipFile
    )

    Write-Verbose "Processing '$($ZipFile.FullName)'"

    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($ZipFile.FullName)

        foreach ($entry in $archive.Entries) {
            $ext = [IO.Path]::GetExtension($entry.FullName).ToLowerInvariant()
            if ($ext -notin '.dwg','.xls','.xlsx','.zip') { continue }

            # Keep the internal folder structure intact.
            $destination = Join-Path $ZipFile.DirectoryName $entry.FullName
            $destinationDir = [IO.Path]::GetDirectoryName($destination)
            if (-not (Test-Path $destinationDir)) {
                [IO.Directory]::CreateDirectory($destinationDir) | Out-Null
            }

            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destination, $true)
        }
    }
    finally {
        # This *always* runs – even if an exception occurred.
        if ($archive) { $archive.Dispose() }
    }
}

while ($true) {
    $zips = Get-ChildItem -Path $SourcePath -Filter *.zip -Recurse -File
    if ($zips.Count -eq 0) {
        Write-Verbose "No more ZIP files found - finishing."
        break
    }

    Write-Verbose ("Found {0} ZIP file{1}" -f $zips.Count, $(if($zips.Count -eq 1) {''} else {'s'}))

    foreach ($zip in $zips) {
        try {
            Expand-ArchiveRecursively -ZipFile $zip
            Remove-Item -LiteralPath $zip.FullName -Force
        }
        catch {
            Write-Warning "Failed to process '$($zip.FullName)': $($_.Exception.Message)"
            $target = Join-Path $CorruptedPath $zip.Name
            Move-Item -LiteralPath $zip.FullName -Destination $target -Force -ErrorAction SilentlyContinue
        }
    }
}