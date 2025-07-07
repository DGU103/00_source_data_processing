<#
.SYNOPSIS
    Recursively unpack *.zip archives (including zips in zips), flattening any
    internal folder structure so that all extracted payload files are placed
    in the same directory where the archive lived. Unprocessable zips are
    moved to a “Corrupted_files” quarantine.
#>

param (
    [Parameter(Mandatory=$true)]
    [ValidateSet(11,12,13)]
    [string] $epc
)

$SourcePath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\indexing\EPC"+$epc+"_Source"
$CorruptedPath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Corrupted_files"

Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

if (-not (Test-Path $CorruptedPath)) {
    New-Item -ItemType Directory -Path $CorruptedPath | Out-Null
}

function Get-UniquePath {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return $Path }

    $dir = [IO.Path]::GetDirectoryName($Path)
    $base = [IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [IO.Path]::GetExtension($Path)
    $i = 1
    do {
        $new = Join-Path $dir ("{0}({1}){2}" -f $base, $i, $ext)
        if (-not (Test-Path $new)) { return $new }
        $i++
    } while ($true)
}

function Expand-ArchiveRecursively {
    param([IO.FileInfo]$ZipFile)

    Write-Verbose "Processing '$($ZipFile.Name)'"

    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($ZipFile.FullName)

        foreach ($entry in $archive.Entries) {

            # Skip directories inside the archive
            if ([String]::IsNullOrWhiteSpace($entry.Name)) { continue }

            $ext = [IO.Path]::GetExtension($entry.Name).ToLowerInvariant()
            if ($ext -notin '.dwg', '.xls', '.xlsx', '.zip') { continue }

            $flatDestination = Join-Path $ZipFile.DirectoryName $entry.Name
            $destPath = Get-UniquePath $flatDestination

            # Extract the entry to a flat file beside the current zip
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $false)

            # If we just wrote another zip, recurse into it, then delete it
            if ($ext -eq '.zip') {
                try {
                    Expand-ArchiveRecursively -ZipFile (Get-Item -LiteralPath $destPath)
                    Remove-Item -LiteralPath $destPath -Force
                }
                catch {
                    # Leave the nested zip in place if it cannot be handled; the main loop will quarantine it
                    Write-Warning "Nested ZIP '$destPath' could not be processed: $($_.Exception.Message)"
                }
            }
        }
    }
    finally {
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

# Optional: post action statistics
$PdfCount = [IO.Directory]::GetFiles($SourcePath, '*.pdf', 'AllDirectories').Count
Write-Verbose "Total PDF files in tree: $PdfCount"
    

