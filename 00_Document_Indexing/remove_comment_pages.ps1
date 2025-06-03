# Load the iTextSharp library
Add-Type -Path "W:\Appli\DigitalAsset\MP\RUYA_data\LocalRepo\00_source_data_processing\00_Document_Indexing\lib\itextsharp.dll"

# Define the function to remove A1 size pages
function Remove-A1PagesFromPDF {
    param (
        [string]$inputFilePath,
        [string]$outputFilePath
    )

    # Open the PDF document
    $reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $inputFilePath
    $document = New-Object iTextSharp.text.Document
    $writer = [iTextSharp.text.pdf.PdfWriter]::GetInstance($document, [System.IO.File]::Create($outputFilePath))
    $document.Open()
    $cb = $writer.DirectContent

    # Iterate through the pages

    for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
         $pageSize = $reader.GetPageSizeWithRotation($i)
         $pageSize
         $width = $pageSize.Width
         $height = $pageSize.Height
         $rotation = $reader.GetPageRotation($i)
        
         # Check if the page size is A1 (594 x 841 mm)
        if (($width -eq 1190.4 -and $height -eq 841.68)) {
            CONTINUE
        }
            else{
            $document.SetPageSize($pageSize)
            $document.NewPage()
            $importedPage = $writer.GetImportedPage($reader, $i)
            
            # Apply the original rotation
            switch ($rotation) {
                90 { $cb.AddTemplate($importedPage, 0, -1, 1, 0, 0, $height) }
                180 { $cb.AddTemplate($importedPage, -1, 0, 0, -1, $width, $height) }
                270 { $cb.AddTemplate($importedPage, 0, 1, -1, 0, $width, 0) }
                default { $cb.AddTemplate($importedPage, 0, 0) }
            }
        }
    }
        

    # Close the document
    $document.Close()
    $reader.Close()
}

# Example usage
$inputFilePath = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\CPPR1-MDM5-ASBJY-08-120001-000F\02\CPPR1-MDM5-ASBJY-08-120001-000F.pdf"
$outputFilePath = "C:\Users\mch107\Downloads\output\output.pdf"
Remove-A1PagesFromPDF -inputFilePath $inputFilePath -outputFilePath $outputFilePath
# $pdf_files = Get-ChildItem -Path "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source" -Recurse -Filter *.pdf