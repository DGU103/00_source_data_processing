param (
    [Parameter(Mandatory=$true)]
    [ValidateSet('11','12','13')]
    [String]$epc
)

$global:scriptname = $MyInvocation.MyCommand.Name -replace '\.ps1$',''
$global:method = "AIM"
$finished = $false

. "$PSScriptRoot\..\Common_Functions.ps1"

Write-Log -Level INFO -Message "====================================="
Write-Log -Level INFO -Message "Running $scriptname for EPCIC $epc"
Write-Log -Level INFO -Message "User: $env:userDomain\$env:UserName"


if($epc -in @('11','12','13')){
    $tag_path =  "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPCIC"+ [string]$epc +"_indexing_report.csv"
}
if (-not (Test-Path -Path $tag_path)) {

    Write-Log -Level ERROR -Message "Path $tag_path not found"
    throw
}

$mtr_path = "\\Qamv3-sapp243\gdp\GDP_StagingArea\MP\MTR\MTR_EPCIC" + [string]$epc + "_Tag_Load_to_AIM-A.csv"

if (-not (Test-Path -Path $tag_path)) {

    Write-Log -Level ERROR -Message "Path $mtr_path not found"
    throw
}

#AIM Export
# $aim_report = "\\qamv3-sapp243\GDP\GDP_StagingArea\MP\Documents\Tag2Doc\EPCIC"+  [string]$epc +"_tag2doc_report.csv"
$aim_report = "\\QAMV3-SFIL102\Home\DGU103\My Documents\Artifacts\AIMA\"+  [string]$epc +"TEST_tag2doc_report.csv"

Write-Log -Level INFO -Message "Build a HashSet of valid tags from MTR"
# 1) Build a HashSet of valid tags from MTR
$mtrData = Import-Csv $mtr_path
$tagSet = [System.Collections.Generic.HashSet[string]]::new()
foreach ($row in $mtrData) {
    $null = $tagSet.Add($row.Tag_number)
}

Write-Log -Level INFO -Message "Detecting Dupicates in Indexing"
# 2) Create a HashSet for detecting duplicates in Indexing
$indexDupes = [System.Collections.Generic.HashSet[string]]::new()

$results = foreach ($row in Import-Csv $tag_path) {

    if ($row.validation_state -eq 'True' -and $tagSet.Contains($row.Tag_number)) {

        # Build a key to detect duplicates
        $key = "$($row.Tag_number)|$($row.Document_number)"

        if (-not $indexDupes.Contains($key)) {
            # We haven't seen this combination yet, so add it
            $null = $indexDupes.Add($key)

            # Output a new object with renamed columns
                [PSCustomObject]@{
                EPCIC = $epc
                Reference_ID = $row.Tag_number
                Document_ID = $row.Document_number
            }
        }
    }
}

$connString = 'Server=QA-SQL-TEST2019;Database=AIM_DEV;Integrated Security=SSPI'

# -------------------------------------------------------------------
# STEP 1 – push batch as Import via TVP
# -------------------------------------------------------------------
$map = [ordered]@{
    EPCIC = 'EPCIC'
    Reference_ID = 'Reference_ID'
    Document_ID = 'Document_ID'
}
# $dtImport = ConvertTo-DataTable -Objects $results -ColumnMap $map
$dtImport = ConvertTo-DataTable $results

Send-Tvp -ConnectionString $connString -Procedure 'dbo.usp_Tag2Doc_Load' -TvpParamName 'NewRows' -TvpTypeName 'dbo.Tag2DocInput' -DataTable $dtImport -ScalarParams @{ ArchiveCurrentIfMissing = 1 }

# -----------------------------------------------------------
# STEP 2 – export the identical data to CSV with alt. columns
# -----------------------------------------------------------
$results | Select-Object @{n='Reference_ID';e={$_.Reference_ID}},
                         @{n='Document_ID';e={$_.Document_ID}},
                         @{n='Action';e={$null}} |
         Export-Csv $aim_report -NoTypeInformation

# -----------------------------------------------------------
# STEP 3 – re-load CSV, flip Import ➜ Current
# -----------------------------------------------------------
$csv = Import-Csv $csvPath

$mapcsv = [ordered]@{
    EPCIC        = 'EPCIC'
    Reference_ID = 'Reference_ID'
    Document_ID  = 'Document_ID'
}
$dtCsv = ConvertTo-DataTable_2 -Objects $csv -ColumnMap $mapcsv

Send-Tvp -ConnectionString $connString `
        -Procedure 'dbo.usp_Tag2Doc_Load' `
        -TvpParamName 'NewRows' `
        -TvpTypeName 'dbo.Tag2DocInput' `
        -DataTable $dtCsv `
        -ScalarParams @{ ArchiveCurrentIfMissing = 1 }


######          OLD     ###############################
# Write-Log -Level INFO -Message "Streaming results into SQL"

# $dataTable = ConvertTo-DataTable $results

# #Streaming to SQL
# $connectionString = "Server=QA-SQL-TEST2019;Database=AIM_DEV;Integrated Security=SSPI"
# $bulk = New-Object System.Data.SqlClient.SqlBulkCopy $connectionString
# $bulk.DestinationTableName = 'dbo.Tag2Doc'

# $bulk.ColumnMappings.Add('EPCIC','EPCIC') | Out-Null
# $bulk.ColumnMappings.Add('Reference_ID','Reference_ID') | Out-Null
# $bulk.ColumnMappings.Add('Document_ID','Document_ID') | Out-Null

# $bulk.BatchSize = 5000
# $bulk.NotifyAfter = 5000
# $bulk.WriteToServer($dataTable)
# $bulk.Close()

# Write-Log -Level INFO -Message "Bulk upload finished. Current state: $1count records inserted"


# Write-Log -Level INFO -Message "Exporting DoctoTag results into CSV."
# $results | Export-Csv -Path $aim_report -NoTypeInformation -force -Delimiter ';'
# Write-Log -Level INFO -Message "Export of DocToTag report is Finished." 

# Write-Log -Level INFO -Message "Uploading results into SQL (current load)"

# $Batch = Import-Csv $aim_report
# Invoke-Tag2DocUpsert -epc $epc -Batch $Batch

# $finished = $true

# Write-Log -Level INFO -Message "Bulk upload finished. Current state: $2count records inserted" -finished $finished



