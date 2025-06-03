# Import the necessary module
Import-Module C:\Users\mch107\Downloads\importexcel.7.4.1\ImportExcel.psd1

function Compare-ExcelFiles {
    param (
        [string]$sourceFolder,
        [string]$targetFolder
    )

    # Get list of files in source and target folders
    $sourceFiles = Get-ChildItem -Path $sourceFolder -Filter *.xlsx
    $targetFiles = Get-ChildItem -Path $targetFolder -Filter *.xlsx

    foreach ($sourceFile in $sourceFiles) {
        $targetFile = $targetFiles | Where-Object { $_.Name -eq $sourceFile.Name }
        if ($targetFile) {
            # Load workbooks
            $sourceWorkbook = Open-ExcelPackage -Path $sourceFile.FullName
            $targetWorkbook = Open-ExcelPackage -Path $targetFile.FullName

            foreach ($sheet in $sourceWorkbook.Workbook.Worksheets) {
                $sourceSheet = $sheet
                $targetSheet = $targetWorkbook.Workbook.Worksheets[$sheet.Name]

                for ($row = 1; $row -le $sourceSheet.Dimension.End.Row; $row++) {
                    for ($col = 1; $col -le $sourceSheet.Dimension.End.Column; $col++) {
                        $sourceCell = $sourceSheet.Cells[$row, $col]
                        $targetCell = $targetSheet.Cells[$row, $col]

                        $sourceValue = $sourceCell.Value
                        $targetValue = $targetCell.Value

                        # Initialize variables for numeric comparison
                        $sourceNum = $null
                        $targetNum = $null

                        if ([double]::TryParse($sourceValue, [ref]$sourceNum) -and [double]::TryParse($targetValue, [ref]$targetNum)) {
                            if ([math]::Abs($sourceNum - $targetNum) / [math]::Abs($sourceNum) -gt 0.005) {
                                $targetCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $targetCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                            }
                        } else {
                            if ($sourceValue -ne $targetValue) {
                                $targetCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $targetCell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                            }
                        }
                    }
                }
            }

            # Save the modified target workbook
            Close-ExcelPackage -ExcelPackage $targetWorkbook -Path $targetFile.FullName
        }
    }
}


# Example usage
$sourceFolderPath = 'C:\Users\mch107\Downloads\ds_source'
$targetFolderPath = 'C:\Users\mch107\Downloads\ds_target'
Compare-ExcelFiles -sourceFolder $sourceFolderPath -targetFolder $targetFolderPath
