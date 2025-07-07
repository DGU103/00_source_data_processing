Set-Location $PSScriptRoot
Clear-Host

class TagObject {
    [String] $Name
    [String] $TYPE
    [String] $SOURCE
    # [String] $DESC
    [String] $DATE
    [String] $PACKAGE
    # [string] $namingtemplate
}

. "$PSScriptRoot\..\Common_Functions.ps1"

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "E3D"
$finished = $false

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"
Write-Log -Level INFO -Message "Running $($MyInvocation.MyCommand.Name). Please Wait"
Write-Log -Level INFO -Message "====================================="

$source_csv_dir = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\E3D\Tagged"
$regex_config_path = '\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\06_Regexp_configs\Full_regex.csv'
$full_regexes = Import-Csv -Delimiter ";" -Path $regex_config_path

# Pre-compile regex patterns
$compiledRegexes = $full_regexes | ForEach-Object {
    [PSCustomObject]@{
        Pattern = [regex]::new("^$($_.Regexp)$")
        Template = $_.Naming_template_ID
    }
}

$csv_files = Get-ChildItem -Path $source_csv_dir -Filter "*-Tagged.csv"

foreach ($file in $csv_files) {
    Write-Log -Level INFO -Message "Processing file: $($file.Name)"
    $records = Import-Csv -Delimiter ";" -Path $file.FullName
    $result = @()

    foreach ($record in $records) {
        if ($record.TYPE -in ("ZONE", "SITE", "NOZZ")) { continue }

        foreach ($regex in $compiledRegexes) {
            if ($regex.Pattern.IsMatch($record.Name)) {
                $tag = New-Object TagObject
                $tag.Name = $record.Name
                $tag.TYPE = $record.TYPE
                $tag.SOURCE = $record.SOURCE
                $tag.DATE = $record.DATE
                $tag.PACKAGE = $record.PACKAGE
                $result += $tag
                break
            }
        }
    }

    if ($file.BaseName -like "EPCIC11*") {

        $outputFile = Join-Path (Join-Path $source_csv_dir "EPC11") "$($file.BaseName)_processed.csv"
    }

    else {$outputFile = Join-Path $source_csv_dir "$($file.BaseName)_processed.csv"}

    try {
        if ($result.Count -gt 0) {
            $result |
                # Sort-Object -Property Name -Unique |
                # Select-Object -Property NAME, TYPE, DATE, NAMINGTEMPLATE |
                Export-Csv -Path $outputFile -NoTypeInformation -Force -Encoding UTF8

            Write-Log -Level INFO -Message "Exported: $outputFile"
        } else {
            Write-Log -Level WARNING -Message "No matching records found in $($file.Name)"
        }
    }
    catch {
        Write-Log -Level ERROR -Message "Failed to export $outputFile. Error: $($_.Exception.Message)"
    }
}

$finished = $true

Write-Log -Level INFO -Message "All TAG exports finished." -finished $finished
