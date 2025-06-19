function Write-Log {
    param(
        #[Parameter(Mandatory=$true)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,
        [switch]$NoTimeStamp,

        [bool]$finished
    )

    $timestamp = if ($NoTimeStamp) { "" } else { "$(Get-Date -Format 'MM-dd-yyyy HH:mm:ss') " }
    $logformat = Get-Date -Format 'MM-dd-yyyy'
    $logLine = "$($timestamp)[$Level] $Message"

    $logpath = "\\als.local\NOC\Data\Appli\DigitalAsset\MP\RUYA_data\Logs\Methods\$method\$logformat"
    # $logpath = "\\QAMV3-SFIL102\Home\DGU103\My Documents\logs\$method\$logformat"

        if(!(test-path -PathType container $logpath)) {
                New-Item -ItemType Directory -Path $logpath
        }

    $logname = "EPCIC$epc" + '_' + $scriptname + '_' + $logformat + '.log'
    $logFile = "$logpath\$logname"

    if (($Level -eq 'DEBUG') -and (-not $global:DEBUG_ENABLED)) {
        return
    }

    # Print to console
    switch ($Level) {
        "INFO" { Write-Host $logLine -ForegroundColor Green }
        "WARN" { Write-Host $logLine -ForegroundColor Yellow }
        "ERROR" { Write-Host $logLine -ForegroundColor Red }
        "DEBUG" { Write-Host $logLine -ForegroundColor Gray }
    }

        # Append to log file
        Add-Content -Path $logFile -Value $logLine


    if ($finished) {Rename-Item -Path $logFile -NewName "OK_$logname" -ErrorAction SilentlyContinue}

    if ($Level -eq 'ERROR') { Rename-Item -Path $logFile -NewName "ERR_$logname" -ErrorAction SilentlyContinue}

}

function Get-Bookmarks {
    param (
        [array]$Bookmarks,
        [array]$Light_Regex,
        [string]$fileBaseName,
        [string]$date,
        [string]$revision_date,
        [string]$reasonText,
        [string]$fileFullName,
        [string]$doctype,
        [string]$doctitle,
        [string]$issuance_code,
        [ref]$tags
    )

    foreach ($bookmark in $Bookmarks) {
        if ($bookmark.Kids) {
            Get-Bookmarks -Bookmarks $bookmark.Kids -Light_Regex $Light_Regex -fileBaseName $fileBaseName -date $date -revision_date $revision_date -reasonText $reasonText -fileFullName $fileFullName -doctype $doctype -doctitle $doctitle -issuance_code $issuance_code -tags $tags
        } else {
            foreach ($regex in $Light_Regex) {
                if ($bookmark.Title.Split()[0] -match $regex.Regexp) {
                    $record = [Tag2Doc]::new()
                    $record.Tag_number = $bookmark.Title.Split()[0]
                    $record.Document_number = $fileBaseName
                    $record.doctitle = $doctitle
                    $record.doctype = $doctype
                    $record.issuance_code = $issuance_code                                        
                    $record.ST = "From bookmarks"
                    $record.DATE = $date
                    $record.doc_date = $revision_date
                    $record.issue_reason = $reasonText
                    $record.file_full_path = $fileFullName
                    $tags.Value += $record
                    break
                }
            }
        }
    }
}
       
function Update-SqlTagdoc {

    param (
        [int]$epc,
        [string]$aim_report
    )
#defining a connection string for SQL server

$connString = "Server=QA-SQL-TEST2019; Database=AIM_DEV; Integrated Security=True;"
$connection = New-Object System.Data.SqlClient.SqlConnection $connString
$connection.Open()

# Read CSV into a DataTable
$dataTable = New-Object System.Data.DataTable
# Define DataTable columns to match Tag2Doc table structure
[void]$dataTable.Columns.Add("Tag2DocID", [int])
[void]$dataTable.Columns.Add("EPCIC", [int])
[void]$dataTable.Columns.Add("Reference_ID", [string])
[void]$dataTable.Columns.Add("Document_ID", [string])


Import-Csv $aim_report -Delimiter ';' | ForEach-Object {
    $row = $dataTable.NewRow()
    $row["Tag2DocID"] = "1"
    $row["EPCIC"] = $epc
    $row["Reference_ID"] = $_.Reference_ID
    $row["Document_ID"] = $_.Document_ID
    $dataTable.Rows.Add($row)
}

$bulk = New-Object System.Data.SqlClient.SqlBulkCopy($connection)
$bulk.DestinationTableName = "Tag2Doc"

try {
    $bulk.WriteToServer($dataTable)
    Write-Host "Bulk insert complete. Rows inserted: " $dataTable.Rows.Count
    $2count = $dataTable.Rows.Count
}
catch {
    Write-Host "Bulk insert failed: $($_.Exception.Message)"
}
finally {
    $bulk.Close()
    $connection.Close()
}

}

function Invoke-Tag2DocUpsert {
    param([System.Collections.IEnumerable]$Batch,
    [String]$epc)

$data = New-Object System.Data.DataTable
# Define DataTable columns to match Tag2Doc table structure
[void]$data.Columns.Add("EPCIC", [int])
[void]$data.Columns.Add("Reference_ID", [string])
[void]$data.Columns.Add("Document_ID", [string])


foreach ($r in $Batch) {
    $row = $data.NewRow()
    $row.EPCIC = $epc
    $row.Reference_ID = $r.Reference_ID
    $row.Document_ID = $r.Document_ID
    $data.Rows.Add($row)
}

# Call the proc with a TVP
$conn = New-Object System.Data.SqlClient.SqlConnection `
          "Server=QA-SQL-TEST2019;Database=AIM_DEV;Integrated Security=SSPI"
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandType = [System.Data.CommandType]::StoredProcedure
$cmd.CommandText = "dbo.usp_Tag2Doc_Load"

$param = $cmd.Parameters.Add("@NewRows",
            [System.Data.SqlDbType]::Structured)
$param.TypeName = "dbo.Tag2DocInput"
$param.Value = $data

$cmd.ExecuteNonQuery()
$conn.Close()

}

function E3D {


    param (
        [bool]$fullrun,
        [bool]$packingvoke,
        [string]$e3d_filters,
        [string]$e3d_links,
        [string]$e3d_tags,
        [string]$e3d_model,    
        [string]$epc  
        )

    if ($e3d_tags -eq 'y') {  
       
        & "$PSScriptRoot\02_E3D\00.01_Export_Tags_from_E3D.ps1"
        & "$PSScriptRoot\02_E3D\02.01_Export_parent_links_from_E3D.ps1"
   
    }
       
   
    $scripts = @("01.00_E3D_Tagged_Item_full_regex_Regex_Filtering.ps1",
    "01.01_E3D_Non-Tagged_Items_Filtering.ps1")

    switch ($e3d_filters) {

        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\02_E3D\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\02_E3D\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\02_E3D\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\02_E3D\$script" -epc 11
            & "$PSScriptRoot\02_E3D\$script" -epc 12
            & "$PSScriptRoot\02_E3D\$script" -epc 13  

            }

    }

    }

    switch ($e3d_links) {

        "11" { & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 11
            & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 12
            & "$PSScriptRoot\02_E3D\02_AIM_3D_model_links.ps1" -epc 13

        }

    }
   
    if ($e3d_model -eq 'y') { & "$PSScriptRoot\02_E3D\03.01_Export_3D_model_from_E3D.ps1" }
   
    if ($fullrun -or $epc) {

        & "$PSScriptRoot\02_E3D\00.01_Export_Tags_from_E3D.ps1"

        $scripts = @("02.01_Export_parent_links_from_E3D.ps1",
                    "01.00_E3D_Tagged_Item_full_regex_Regex_Filtering.ps1",
                    "01.01_E3D_Non-Tagged_Items_Filtering.ps1",
                    "02_AIM_3D_model_links.ps1")
       
        foreach ($script in $scripts) {

            if ($packingvoke) {& "$PSScriptRoot\02_E3D\$script" -epc $epc}

                else {

                & "$PSScriptRoot\02_E3D\$script" -epc 11
                & "$PSScriptRoot\02_E3D\$script" -epc 12
                & "$PSScriptRoot\02_E3D\$script" -epc 13

                }
       
            }

        & "$PSScriptRoot\02_E3D\03.01_Export_3D_model_from_E3D.ps1"

    }

}

function E_I {

    param (
        [bool]$fullrun,    
        [bool]$packingvoke,    
        [string]$EI_tags,
        [string]$EI_props,
        [string]$epc
    )

    <# Tag export #>

    switch ($EI_tags) {

        "11" { & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 11  
            & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 12                            
            & "$PSScriptRoot\03_EI\00_Extract_Tags_From_EI.ps1" -epc 13
        }        

    }
 
    <# Props export #>  

    switch ($EI_props) {

        "11" { & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 11  
            & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 12                            
            & "$PSScriptRoot\03_EI\00_Extract_E&I_Properties_From_EI.ps1" -epc 13
        }        

    }
   
    if ($fullrun -or $epc) {

        $scripts = @("00_Extract_Tags_From_EI.ps1",
        "00_Extract_E&I_Properties_From_EI.ps1")

foreach ($script in $scripts) {

    if ($packingvoke) {& "$PSScriptRoot\03_EI\$script" -epc $epc}

    else {

        & "$PSScriptRoot\03_EI\$script" -epc 11
        & "$PSScriptRoot\03_EI\$script" -epc 12
        & "$PSScriptRoot\03_EI\$script" -epc 13

            }

        }

    }
}

function Engineering {

    param (
        [bool]$fullrun,      
        [bool]$packinvoke,      
        [string]$Eng_tags,
        [string]$Eng_props,
        [string]$epc
    )

    <# Tag export #>
 
    switch ($Eng_tags) {

        "11" { & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 11  
            & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 12                            
            & "$PSScriptRoot\04_Engineering\01.01_Export_Tags_From_Engineering.ps1" -epc 13
        }        

    }

    <# Properties export #>  

    switch ($Eng_props) {

        "11" { & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 11 }

        "12" { & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 12 }

        "13" { & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 13 }

        "all" {

            & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 11  
            & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 12                            
            & "$PSScriptRoot\04_Engineering\01.02_Export_Properties_from_Engineering.ps1" -epc 13
        }        

    }

    if ($fullrun -or $packingvoke) {

        $scripts = @("01.01_Export_Tags_From_Engineering.ps1",
                    "01.02_Export_Properties_from_Engineering.ps1")
 
        foreach ($script in $scripts) {

            if ($packingvoke) {& "$PSScriptRoot\04_Engineering\$script" -epc $epc}

            else {

            & "$PSScriptRoot\04_Engineering\$script" -epc 11
            & "$PSScriptRoot\04_Engineering\$script" -epc 12
            & "$PSScriptRoot\04_Engineering\$script" -epc 13

        }

        }

    }
}

function Diagrams {

    param (
        [bool]$fullrun,      
        [string]$Dia_tags
    )
   

    $scripts = @("01_Extract_Tags_from_Diagrams.ps1",
    "02_Extract_SCGROU_for_SVG_export.ps1")

    if ($Dia_tags -eq 'y') { foreach ($script in $scripts) { & "$PSScriptRoot\01_Diagrams\$script" } }
   
    if ($fullrun){ foreach ($script in $scripts) { & "$PSScriptRoot\01_Diagrams\$script" } }
 
}    
function Indexing {

    param (
        [bool]$fullrun,        
        [bool]$packingvoke,        
        [string]$meta_update,
        [string]$epc_envoke,
        [string]$aim_index,
        [string]$epc
    )
   
    #Data Fetching from MANASA

    $scripts = @("01.00_Delete_metadata_from_folder.ps1",
    "01.01_Extract_metadata_from_DMS.ps1",
    "01.02_Process_Metadata.ps1",
    "01.03_Extract_PDFs_from_MANASA.ps1")

    switch ($meta_update) {

        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13  

            }

        }    

    }

    <#Pure Indexing #>

    $scripts = @("01.04_Doc_indexing_multyTread.ps1",
    "01.05_Indexing_result_postProcessing.ps1")

    switch ($epc_envoke) {

        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13  

            }

        }        

    }

    <# AIM-A Section #>

    $scripts = @("02.01_Document_Register_for_AIM.ps1",
    "02.02_Publish_Doc_to_Tag.ps1",
    "02.03_PDF_copy_to_AIM.ps1")

    switch ($aim_index) {

       
        "11" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 11} }

        "12" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 12} }

        "13" { foreach ($script in $scripts) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc 13} }

        "all" {

            foreach ($script in $scripts) {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13  

            }

        }

    }

    if ($fullrun -or $packingvoke) {
       
        $scripts = @("01.00_Delete_metadata_from_folder.ps1",
            "01.01_Extract_metadata_from_DMS.ps1",
            "01.02_Process_Metadata.ps1",
            "01.03_Extract_PDFs_from_MANASA.ps1",
            "01.04_Doc_indexing_multyTread.ps1",
            "01.05_Indexing_result_postProcessing.ps1",
            "02.01_Document_Register_for_AIM.ps1",
            "02.02_Publish_Doc_to_Tag.ps1",
            "02.03_PDF_copy_to_AIM.ps1")

        foreach ($script in $scripts) {

            if ($packinvoke) {& "$PSScriptRoot\00_Document_Indexing\$script" -epc $epc}

            else {

            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 11
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 12
            & "$PSScriptRoot\00_Document_Indexing\$script" -epc 13
           
            }

        }
    }
}