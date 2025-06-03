function Export-MergedExcelDataToCsv {
    param (
        [Parameter(Mandatory = $true)]
        [string]$RootFolder,

        [Parameter(Mandatory = $true)]
        [string]$OutputFolder
    )

    $outputCsv = Join-Path $OutputFolder "master_output.csv"

    # Header keywords including "Line Number"
    $headerKeywords = @("Equipment No", "EquipmentNo", "Tag No", "TagNo", "Tag Number", "TagNumber", "Line Number")
    

    Write-Host "[INFO] Scanning for Excel files in: $RootFolder"
    $rowexcelFiles = Get-ChildItem -Path $RootFolder -Recurse -Include *.xls, *.xlsx |
        Where-Object { $_.Name -notmatch 'CRS' }
    $excelFiles = @()
    foreach ($file in $rowexcelFiles) {
        $file.Directory
        $xml_file_path = Get-ChildItem -Path $file.Directory -Filter *.xml | Select-Object -First 1
        [xml]$XmlDocument = Get-Content $xml_file_path
        $XmlNamespace = @{ a = "http://www.aveva.com/VNET/eiwm" }
        $doctype = (Select-Xml -Xml $XmlDocument -XPath "//a:Template/a:Object/a:Characteristic[a:Name='pjc_doc_type']/a:Value" -Namespace $XmlNamespace).Node.InnerText
        if ($doctype -match "LIS|REG|LST") {
            $excelFiles += $file
        }
    }
    $totalFiles = $excelFiles.Count
    Write-Host "[INFO] Found $totalFiles Excel file(s) to process."

    if ($totalFiles -eq 0) {
        Write-Host "[INFO] No files to process. Exiting."
        return
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false


    $fileIndex = 0
    foreach ($file in $excelFiles) {
        $masterRows = @()
        $fileIndex++
        Write-Host "[INFO] Processing file $fileIndex of $totalFiles : $($file.FullName)"
        $workbook = $excel.Workbooks.Open($file.FullName)

        foreach ($worksheet in $workbook.Sheets) {
            $usedRange = $worksheet.UsedRange
            $rowCount = [Math]::Min($usedRange.Rows.Count, 20)
            $colCount = [Math]::Min($usedRange.Columns.Count, 20)

            $headerRow = $null
            $idCol = $null
            for ($row = 1; $row -le $rowCount; $row++) {
                for ($col = 1; $col -le $colCount; $col++) {
                    $cellValue = $usedRange.Cells.Item($row, $col).Text.Trim()
                    if ($headerKeywords -contains $cellValue) {
                        $headerRow = $row
                        $idCol = $col
                        break
                    }
                }
                if ($headerRow) { break }
            }

            if (-not $headerRow -or -not $idCol) {
                Write-Host "[WARN] Header row with ID column not found in sheet '$($worksheet.Name)'. Skipping..."
                continue
            }

            # Read headers from the full width of the sheet
            $fullColCount = $usedRange.Columns.Count
            # $headers = @{}
            # for ($col = 1; $col -le $fullColCount; $col++) {
            #     $header = $usedRange.Cells.Item($headerRow, $col).Text
            #     $headers[$col] = if ($header) { $header } else { "Column$col" }
            # }

            $fullRowCount = $usedRange.Rows.Count
            for ($row = $headerRow + 1; $row -le $fullRowCount; $row++) {
                $idValue = $usedRange.Cells.Item($row, $idCol).Text
                $masterRows += $idValue
                # if ([string]::IsNullOrWhiteSpace($idValue)) { $idValue = "Unknown" }

                # for ($col = 1; $col -le $fullColCount; $col++) {
                #     $value = $usedRange.Cells.Item($row, $col).Text
                #     if (-not [string]::IsNullOrWhiteSpace($value)) {
                #         $columnName = $headers[$col]
                #         $masterRows += [PSCustomObject]@{
                #             'Tag Number'      = $idValue
                #             'Attribute Name'  = $columnName
                #             'Attribute Value' = $value
                #         }
                #     }
                # }
            }
        }

        $workbook.Close($false)
        $outputCsv = Join-Path $OutputFolder ($file.Basename + '.csv')
        Write-Host "[INFO] Exporting merged data to: $outputCsv"
        $masterRows | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
        Write-Host "[SUCCESS] Master CSV exported successfully."
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # if ($masterRows.Count -gt 0) {
    #     Write-Host "[INFO] Exporting merged data to: $outputCsv"
    #     $masterRows | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
    #     Write-Host "[SUCCESS] Master CSV exported successfully."
    # } else {
    #     Write-Host "[INFO] No data found to export."
    # }
}
Export-MergedExcelDataToCsv -RootFolder "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source" `
                            -OutputFolder "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\Temp"
