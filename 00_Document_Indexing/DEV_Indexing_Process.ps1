
<#--------------------------------------------------------------------
    DEV_Indexing.ps1
    Triggers PDF, Excel Processing.
    Supported parameters: -pdf , -excel
    TO DO: DWG Processing , Speed Improvements

CAUTION: DEV. Results may vary
--------------------------------------------------------------------#>

param (
    [Parameter(Mandatory=$true)]
    [string]$epc,
    [switch]$EnableDebug,
    [switch]$pdf,
    [switch]$excel
)

if ($EnableDebug.IsPresent) {
    $global:DEBUG_ENABLED = $true
} else {
    $global:DEBUG_ENABLED = $false
}

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "INDEXING"

#Load common file
. "$PSScriptRoot\..\Common_Functions.ps1"

$global:jobs_list = @() # global collection for aggregation

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "Running $scriptname for EPCIC $epc"
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"

if($global:DEBUG_ENABLED) {
    Write-Log -Level INFO -Message "DEBUG logging is ENABLED."
} else {
    Write-Log -Level INFO -Message "DEBUG logging is DISABLED."
}

# If the caller supplied neither switch → do both
if(-not $pdf -and -not $excel){
    $pdf = $true
    $excel = $true
}

################################################################################
# --- Adjust for Performance Boost ---
################################################################################

$pdfparts = 10
$excelparts = 10

################################################################################
# --- PATHS ---
################################################################################

if($epc -in @('11','12','13')) { $root_path = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\" }
elseif($epc -eq '06') { $root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\CPP03\Source\Indexing\" }
elseif($epc -eq '05') { $root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\" }
else { throw "Unsupported EPC $epc" }

$files_Dir = Join-Path $root_path ("EPC" + $epc + "_Source")


# $tag_report = "H:\My Documents\Artifacts\Indexing\out\EPCIC12_Tagsiki.csv"
# $doc_report = "H:\My Documents\Artifacts\Indexing\out\EPCIC12_Docsiki.csv"


$local_path = $PSScriptRoot
#$files_Dir = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing"
#$tag_report = Join-Path $root_path ("EPCIC" + $epc + "_indexing_report.csv")
#$doc_report = Join-Path $root_path ("EPCIC" + $epc + "DOC_indexing_report.csv")


################################################################################
# --- HELPER INDEXING BATCH ---
################################################################################
function Start-IndexingBatch {
    param(
        [string]$JobName,
        [System.IO.FileInfo[]]$Files,
        [int]$Parts,
        [string]$Epc,
        [string]$LocalPath,
        [ScriptBlock]$Work
    )
    if(-not $Files){ Write-Log -Level WARN -Message "No files for '$JobName'."; return }


[int]$batchSize = [Math]::Ceiling($Files.Count / $Parts)
[int]$offset    = 0
foreach($i in 1..$Parts){

    $end   = [Math]::Min($offset + $batchSize - 1, $Files.Count - 1)
    $slice = $Files[$offset..$end]
    $batch = $slice.FullName

        if($batch.Count -eq 0){ break }

        $jobId = "EPC${Epc}_Batch${i}_${JobName}"
        Write-Log -Level INFO -Message "Start-Job $jobId ($($batch.Count) files)"
        $job = Start-Job -Name $jobId -ArgumentList $batch,$LocalPath,$Epc,$global:DEBUG_ENABLED,$global:scriptname -ScriptBlock $Work
        $global:jobs_list += $job
        $offset += $batchSize
  }

}

################################################################################
# --- PDF PROCESSING  ---
################################################################################


if($pdf){
    Write-Log -Level INFO -Message "Collecting PDF files from: $files_Dir"
    $pdffiles = Get-ChildItem -Path $files_Dir -Recurse -Include *.pdf -File
    Write-Log -Level INFO -Message "Found $($pdffiles.Count) PDFs."

    Start-IndexingBatch -JobName 'PDF_Indexing' -Files $pdffiles -Parts $pdfparts  -Epc $epc -LocalPath $PSScriptRoot -Work {
        param($files,$local_path,$epc,$DEBUG_ENABLED,$inst,$scriptname)

        . "$using:PSScriptRoot\..\Common_Functions.ps1"
        
        Invoke-PDFIndexing -Files $files -Epc $epc -LocalPath $local_path
    }
}

################################################################################
# --- EXCEL PROCESSING  ---
################################################################################


if($excel){
    $files = [System.IO.FileInfo[]]$files
    Write-Log -Level INFO -Message "Collecting EXcel files from: $files_Dir"
    $excelfiles = Get-ChildItem -Path $files_Dir -Recurse -Include *.xls,*.xlsx -File |
                  Where-Object { $_.Name -notmatch 'CRS' }

    Write-Log -Level INFO -Message "Found $($excelfiles.Count) Excels."
    Start-IndexingBatch -JobName 'Excel_Indexing' -Files $excelfiles -Parts $excelparts -Epc $epc -LocalPath $root_path -Work {
        param($files,$local_path,$epc,$DEBUG_ENABLED,$scriptname)
        $files = $files | Foreach-Object {[System.IO.FileInfo]::new($_)}
        $files = [System.IO.FileInfo[]] $files
        . "$using:PSScriptRoot\..\Common_Functions.ps1"
        Invoke-ExcelIndexing -Files $files -Epc $epc -LocalPath $local_path
    }
}

################################################################################
# --- Wait → Merge → Export ---
################################################################################

if($jobs_list.Count -eq 0){
    Write-Log -Level WARN -Message 'Nothing was queued - exiting.'
    return
}

Write-Log -Level INFO -Message 'Waiting for all jobs to finish'
Wait-Job -Job $jobs_list

$tagCollection = @()
$docCollection = @()
$timeCollection = @()

foreach($j in $jobs_list){
    $r = Receive-Job -Job $j
    if($r){
        $tagCollection += $r.Tag2Doc
        $docCollection += $r.Doc2Doc
        $timeCollection += $r.Times
    }
}
Remove-Job -Job $jobs_list -Force

$tagCollection = $tagCollection |
                 Where-Object { -not [string]::IsNullOrEmpty($_.Document_number) } |
                 Select-Object Tag_number, Document_number, doctitle, doctype,
                               issuance_code, ST, DATE, doc_date, issue_reason, file_full_path -Unique

$docCollection = $docCollection |
                 Where-Object { -not [string]::IsNullOrEmpty($_.ref_doc_id) } |
                 Select-Object source_doc_id, ref_doc_id, DATE

Write-Log -Level INFO -Message "Exporting $($tagCollection.Count) TAG records → $tag_report"
try {
    $tagCollection | Export-Csv -Path $tag_report -NoTypeInformation -Encoding UTF8 -Force
}
catch {
    Write-Log -Level ERROR -Message "Failed to export tags CSV - $($_.Exception.Message)"
}

Write-Log -Level INFO -Message "Exporting $($docCollection.Count) DOC records → $doc_report"
try {
    $docCollection | Export-Csv -Path $doc_report -NoTypeInformation -Encoding UTF8 -Force
}
catch {
    Write-Log -Level ERROR -Message "Failed to export docs CSV - $($_.Exception.Message)"
}

if($timeCollection.Count){
    $avg = ($timeCollection | Measure-Object -Average).Average
    Write-Log -Level INFO -Message ("Average processing time per file: {0:N2} s" -f $avg)
}

$finished = $true
Write-Log -Level INFO -Message "Multithreaded indexing complete"