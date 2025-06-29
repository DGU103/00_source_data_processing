param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('05','06','11','12','13','6')]
    [int]$epc
)

class TagNumber {
    [string] $Tag_number
    [bool] $validation_state
    [string] $Document_number
    [string] $doctype
    [string] $doctitle
    [string] $issuance_code
    [string] $ST
    [string] $DATE
    [string] $doc_date
    [string] $issue_reason
    [string] $SourceType
    [string] $discipline
    [string] $Document_Hyper_Link
}

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "INDEXING"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "Running $scriptname for EPCIC $epc"
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"

$Host.UI.RawUI.WindowTitle = "Document indexing postrpocessing for disciplines EPC $epc"
if($epc -in @(11,12,13)){$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\"}
elseif($epc -eq '06' -or $epc -eq '6'){$epc = '06' 
$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\CPP03\Source\Indexing\"}
elseif($epc -eq '05' -or $epc -eq '06'){$epc = '05' 
$root_path = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\WHP03\Source\Indexing\"}

# $hyperlink_path = $root_path + "EPC" + $epc + "Source"
$indexing_report_path = $root_path + "EPCIC" + $epc + "_indexing_report.csv"
$indexing_report = Import-Csv $indexing_report_path

#$regexes = Import-Csv "C:\Users\DGU103\Downloads\GIT\00_source_data_processing\06_Regexp_configs\Full_regex.csv" -Delimiter ";"
$regexes = Import-Csv "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Full_regex.csv" -Delimiter ";"
if (-not $regexes) {
    Write-Log -Level ERROR -Message "Missing config in \\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Full_regex.csv  Error: $($_.Exception.Message)"
    BREAK
}
$report = [TagNumber[]]::new($indexing_report.Count)

$count = $indexing_report.Count
$size = $indexing_report.Count - 1
$index = 0
foreach ($i in 0..$size) {
    $index = $index + 100 / $count
    $progress = [System.Math]::Round($index, 2)
    Write-Progress -Activity "Overall progress:" -Status "$progress% Complete from $count records." -PercentComplete $index
    

    $record = New-Object TagNumber
    $record.Tag_number = $indexing_report[$i].Tag_number
    $record.Document_number = $indexing_report[$i].Document_number
    $record.doctype = $indexing_report[$i].doctype
    $record.doctitle = $indexing_report[$i].doctitle
    $record.issuance_code = $indexing_report[$i].issuance_code
    $record.ST = $indexing_report[$i].ST
    $record.DATE = $indexing_report[$i].DATE
    $record.doc_date = $indexing_report[$i].doc_date
    $record.issue_reason = $indexing_report[$i].issue_reason
    $record.SourceType = $indexing_report[$i].SourceType

    $hyper_link = '=HYPERLINK("' + $indexing_report[$i].file_full_path + '","' + $indexing_report[$i].Document_number +'")'
    $record.Document_Hyper_Link = $hyper_link
    $disc = $false
    foreach ($regex in $regexes) {
        if ($indexing_report[$i].Tag_number -match $regex.Regexp) {
            $record.discipline = $regex.discipline
            $record.validation_state = $true
            $disc = $true
            BREAK
        }
    }
    if (-not $disc) {
        $record.discipline = "Not identified"
        $record.validation_state = $false
    }
    $report[$i]= $record
}

Write-Log -Level INFO -Message "Exporting Postprocessed Results..."
# $report | Select-Object -Unique | Export-Csv -Path $indexing_report_path -NoTypeInformation -Encoding UTF8 -Force
$report | Export-Csv -Path $indexing_report_path -NoTypeInformation -Encoding UTF8 -Force

$finished = $true
Write-Log -Level INFO -Message "PostProcessing Is Fininshed" -finished $finished
