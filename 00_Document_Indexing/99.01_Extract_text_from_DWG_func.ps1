# Load BricsCAD COM object
$bricscadApp = New-Object -ComObject BricscadApp.AcadApplication

# Function to extract text notes from a block recursively
function Get-TextNotesFromBlock {
    param (
        [object]$block
    )

    $textNotesCollection = @()

    # Iterate through each entity in the block
    foreach ($entity in $block) {
        # Check if the entity is a Text or MText object
        if ($entity.ObjectName -eq "AcDbText" -or $entity.ObjectName -eq "AcDbMText") {
            $textNote = [PSCustomObject]@{
                TextValue = $entity.TextString
                XPosition = $entity.InsertionPoint[0]
                YPosition = $entity.InsertionPoint[1]
            }
            $textNotesCollection += $textNote
        }
        # Check if the entity is a BlockReference
        elseif ($entity.ObjectName -eq "AcDbBlockReference") {
            $subBlockRecord = $entity.BlockTableRecord
            if ($subBlockRecord -ne $null) {
                $subBlock = $subBlockRecord.Open()
                if ($subBlock -ne $null) {
                    $textNotesCollection += Get-TextNotesFromBlock -block $subBlock
                }
            }
            # Extract attribute values from BlockReference
            foreach ($attribute in $entity.GetAttributes()) {
                $textNote = [PSCustomObject]@{
                    TextValue = $attribute.TextString
                    XPosition = $attribute.InsertionPoint[0]
                    YPosition = $attribute.InsertionPoint[1]
                }
                $textNotesCollection += $textNote
            }
        }
    }

    return $textNotesCollection
}

# Function to extract text notes from a DWG file
function Get-TextNotesFromDWG {
    param (
        [string]$dwgFilePath
    )

    # Open the DWG file
    $doc = $bricscadApp.Documents.Open($dwgFilePath)
    $textNotesCollection = @()

    # Iterate through each layout in the document
    foreach ($layout in $doc.Layouts) {
        # Get the Block object for the layout
        $block = $layout.Block

        # Get text notes from the block recursively
        $textNotesCollection += Get-TextNotesFromBlock -block $block
    }

    # Iterate through each entity in the ModelSpace to collect text notes not in blocks
    foreach ($entity in $doc.ModelSpace) {
        if ($entity.ObjectName -eq "AcDbText" -or $entity.ObjectName -eq "AcDbMText") {
            $textNote = [PSCustomObject]@{
                TextValue = $entity.TextString
                XPosition = $entity.InsertionPoint[0]
                YPosition = $entity.InsertionPoint[1]
            }
            $textNotesCollection += $textNote
        }
        # Check if the entity is a BlockReference and extract attribute values
        elseif ($entity.ObjectName -eq "AcDbBlockReference") {
            foreach ($attribute in $entity.GetAttributes()) {
                $textNote = [PSCustomObject]@{
                    TextValue = $attribute.TextString
                    XPosition = $attribute.InsertionPoint[0]
                    YPosition = $attribute.InsertionPoint[1]
                }
                $textNotesCollection += $textNote
            }
        }
    }

    # Close the document
    $doc.Close()

    return $textNotesCollection
}

# Function to process all DWG files in a folder
function Process-DWGFilesInFolder {
    param (
        [string]$folderPath
    )

    # Get all DWG files in the folder
    $dwgFiles = Get-ChildItem -Path $folderPath -Filter *.dwg

    foreach ($file in $dwgFiles) {
        Write-Output "Processing file: $($file.FullName)"
        $textNotesCollection = Get-TextNotesFromDWG -dwgFilePath $file.FullName

        # Output the collected text notes to proc_result.txt
        $textNotesCollection | ForEach-Object { 
            "$($_.TextValue) | X: $($_.XPosition) | Y: $($_.YPosition)" 
        } | Out-File -FilePath "$folderpath\proc_result.txt" -Force
    }
}

# Example usage
$folderPath = "W:\Appli\DigitalAsset\MP\RUYA_data\Source\Indexing\EPC13_Source\CPPR1-MDM5-ASBJA-10-R54062-0001\03"
Process-DWGFilesInFolder -folderPath $folderPath 

# Quit BricsCAD application
$bricscadApp.Quit()
